from flask import Flask, render_template, request, send_file, redirect, session, url_for, jsonify
import os
import hmac
import hashlib
import requests
import functools
import openpyxl
from datetime import datetime, timezone
from PIL import Image, ImageDraw, ImageFont
import zipfile
import io
from dotenv import load_dotenv
from supabase import create_client
from postgrest.exceptions import APIError

# Load environment variables from .env file
load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
FLASK_SECRET_KEY = os.getenv("FLASK_SECRET_KEY")

if not FLASK_SECRET_KEY:
    raise RuntimeError(
        "FLASK_SECRET_KEY is not set in your .env file. "
        "Generate one and add it before running the app."
    )

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

app = Flask(__name__)
app.secret_key = FLASK_SECRET_KEY

# Session cookie hardening
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
)

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

FREE_TRIAL_LIMIT = 15

LEMONSQUEEZY_API_KEY = os.getenv("LEMONSQUEEZY_API_KEY")
LEMONSQUEEZY_STORE_ID = os.getenv("LEMONSQUEEZY_STORE_ID")
LEMONSQUEEZY_WEBHOOK_SECRET = os.getenv("LEMONSQUEEZY_WEBHOOK_SECRET")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY")

VARIANT_IDS = {
    'monthly': os.getenv("LEMONSQUEEZY_VARIANT_MONTHLY"),
    'sixmonth': os.getenv("LEMONSQUEEZY_VARIANT_6MONTH"),
    'yearly': os.getenv("LEMONSQUEEZY_VARIANT_YEARLY"),
}

# ---------------------------------------------------------------------------
# Global error handling -- expired sessions redirect cleanly instead of
# showing a raw crash page. This was a recurring issue during testing;
# fixed once here at the app level instead of patched per-route.
# ---------------------------------------------------------------------------

@app.errorhandler(APIError)
def handle_api_error(error):
    error_message = str(error)
    if 'JWT expired' in error_message or 'invalid claim' in error_message.lower():
        session.clear()
        return redirect('/login')
    # Any other unexpected database error: log it, show a generic message
    # rather than a raw stack trace to the user.
    app.logger.error(f"Supabase API error: {error_message}")
    return render_template('login.html', error="Something went wrong. Please log in again."), 500


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_profile_for_user(user_id, access_token):
    """
    Fetch a user's own profile row, querying AS that user (their access token
    is attached to the request) so Row Level Security allows it. Never uses
    the service_role key here -- this always runs with the user's own
    permissions, matching the "Users can view own profile" policy.
    """
    client = create_client(SUPABASE_URL, SUPABASE_KEY)
    client.postgrest.auth(access_token)
    result = client.table('profiles').select('*').eq('id', user_id).single().execute()
    return result.data


def increment_free_certificates_used(user_id, access_token, count):
    """
    Adds `count` to the user's free_certificates_used total. Runs as the
    user themselves (their own access token), matching the "Users can
    update own profile" RLS policy -- never uses the service_role key.
    """
    client = create_client(SUPABASE_URL, SUPABASE_KEY)
    client.postgrest.auth(access_token)
    current = client.table('profiles').select('free_certificates_used').eq('id', user_id).single().execute()
    new_total = (current.data.get('free_certificates_used') or 0) + count
    client.table('profiles').update({'free_certificates_used': new_total}).eq('id', user_id).execute()


def update_last_sign_in(user_id, access_token):
    client = create_client(SUPABASE_URL, SUPABASE_KEY)
    client.postgrest.auth(access_token)
    client.table('profiles').update(
        {'last_sign_in': datetime.now(timezone.utc).isoformat()}
    ).eq('id', user_id).execute()

def get_service_client():
    """
    Returns a Supabase client authenticated with the service_role key.
    Bypasses RLS entirely -- only ever used inside the webhook handler,
    since Lemon Squeezy webhooks have no user session/access token.
    """
    return create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)

def verify_lemonsqueezy_signature(raw_body: bytes, signature_header: str) -> bool:
    if not signature_header:
        return False
    digest = hmac.new(
        LEMONSQUEEZY_WEBHOOK_SECRET.encode('utf-8'),
        raw_body,
        hashlib.sha256
    ).hexdigest()
    return hmac.compare_digest(digest, signature_header)

def login_required(view_func):
    @functools.wraps(view_func)
    def wrapped(*args, **kwargs):
        if not session.get('user_id') or not session.get('access_token'):
            return redirect('/login')
        return view_func(*args, **kwargs)
    return wrapped


def subscription_required(view_func):
    @functools.wraps(view_func)
    def wrapped(*args, **kwargs):
        profile = get_profile_for_user(session['user_id'], session['access_token'])
        if not profile:
            return redirect('/login')

        is_admin = profile.get('role') in ('super_admin', 'sub_admin')
        is_subscribed = profile.get('subscription_status') == 'active'
        trial_remaining = FREE_TRIAL_LIMIT - (profile.get('free_certificates_used') or 0)

        has_access = is_admin or is_subscribed or trial_remaining > 0
        if not has_access:
            return redirect('/subscription-required')

        return view_func(*args, **kwargs)
    return wrapped

@app.route('/subscribe')
@login_required
def subscribe():
    plan = request.args.get('plan')
    if plan not in VARIANT_IDS or not VARIANT_IDS[plan]:
        return "Invalid plan", 400

    variant_id = VARIANT_IDS[plan]
    user_id = session['user_id']
    user_email = session['email']

    url = "https://api.lemonsqueezy.com/v1/checkouts"
    headers = {
        "Authorization": f"Bearer {LEMONSQUEEZY_API_KEY}",
        "Accept": "application/vnd.api+json",
        "Content-Type": "application/vnd.api+json",
    }
    payload = {
        "data": {
            "type": "checkouts",
            "attributes": {
                "checkout_data": {
                    "email": user_email,
                    "custom": {"user_id": user_id}
                }
            },
            "relationships": {
                "store": {"data": {"type": "stores", "id": LEMONSQUEEZY_STORE_ID}},
                "variant": {"data": {"type": "variants", "id": variant_id}}
            }
        }
    }

    r = requests.post(url, json=payload, headers=headers)
    print("LEMONSQUEEZY RESPONSE:", r.status_code, r.text)
    r.raise_for_status()
    checkout_url = r.json()['data']['attributes']['url']

    return redirect(checkout_url)

    #-------------------------------------------------------
    # webhook handler for Lemon Squeezy subscription events. 
    #-------------------------------------------------------
@app.route('/webhooks/lemonsqueezy', methods=['POST'])
def lemonsqueezy_webhook():
    signature = request.headers.get('X-Signature', '')
    raw_body = request.get_data()

    if not verify_lemonsqueezy_signature(raw_body, signature):
        return jsonify({'error': 'invalid signature'}), 401

    event = request.get_json()
    event_name = event['meta']['event_name']
    custom_data = event['meta'].get('custom_data', {})
    data = event['data']
    attributes = data['attributes']
    ls_subscription_id = data['id']

    client = get_service_client()

    # ------------------------------------------------------------------
    # New subscription -- match by the user_id we passed at checkout
    # ------------------------------------------------------------------
    if event_name == 'subscription_created':
        user_id = custom_data.get('user_id')
        if user_id:
            client.table('profiles').update({
                'subscription_status': 'active',
                'ls_subscription_id': ls_subscription_id,
                'ls_customer_id': str(attributes.get('customer_id')),
            }).eq('id', user_id).execute()

    # ------------------------------------------------------------------
    # All other events -- match by the stored ls_subscription_id,
    # since these fire on existing subscriptions, not new ones
    # ------------------------------------------------------------------
    elif event_name in ('subscription_updated', 'subscription_payment_success',
                         'subscription_payment_recovered', 'subscription_resumed',
                         'subscription_unpaused', 'subscription_plan_changed'):
        status = attributes.get('status')
        new_status = 'active' if status in ('active', 'on_trial') else 'inactive'
        client.table('profiles').update({
            'subscription_status': new_status,
        }).eq('ls_subscription_id', ls_subscription_id).execute()

    elif event_name in ('subscription_cancelled', 'subscription_expired',
                         'subscription_paused', 'subscription_payment_failed'):
        client.table('profiles').update({
            'subscription_status': 'inactive',
        }).eq('ls_subscription_id', ls_subscription_id).execute()

    elif event_name == 'dispute_created':
        client.table('profiles').update({
            'subscription_status': 'inactive',
        }).eq('ls_subscription_id', ls_subscription_id).execute()

    elif event_name == 'dispute_resolved':
        status = attributes.get('status')
        if status == 'won':
            client.table('profiles').update({
                'subscription_status': 'active',
            }).eq('ls_subscription_id', ls_subscription_id).execute()

    # customer_updated, subscription_payment_refunded -- logged but no
    # status change required for your current access model
    else:
        app.logger.info(f"Unhandled LS webhook event: {event_name}")

    return jsonify({'received': True}), 200

def admin_required(view_func):
    @functools.wraps(view_func)
    def wrapped(*args, **kwargs):
        profile = get_profile_for_user(session['user_id'], session['access_token'])
        if not profile or profile.get('role') != 'super_admin':
            return redirect('/generate-certificates')
        return view_func(*args, **kwargs)
    return wrapped


# ---------------------------------------------------------------------------
# Public routes
# ---------------------------------------------------------------------------

@app.route('/')
def landing():
    # Public landing page -- no login required. Logged-in users get a
    # "Go to App" button instead of "Sign Up" (checked client-side isn't
    # reliable, so we check session here).
    is_logged_in = bool(session.get('user_id'))
    return render_template('landing.html', is_logged_in=is_logged_in)


# ---------------------------------------------------------------------------
# Auth routes
# ---------------------------------------------------------------------------

@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == 'GET':
        return render_template('signup.html', error=None, success=None)

    email = request.form.get('email', '').strip()
    password = request.form.get('password', '')

    try:
        supabase.auth.sign_up({"email": email, "password": password})
    except Exception as e:
        return render_template('signup.html', error=str(e), success=None)

    return render_template(
        'signup.html',
        error=None,
        success="Account created! Check your email to confirm, then log in."
    )


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        return render_template('login.html', error=None)

    email = request.form.get('email', '').strip()
    password = request.form.get('password', '')

    try:
        result = supabase.auth.sign_in_with_password({"email": email, "password": password})
    except Exception:
        return render_template('login.html', error="Invalid email or password.")

    session['user_id'] = result.user.id
    session['email'] = result.user.email
    session['access_token'] = result.session.access_token

    update_last_sign_in(result.user.id, result.session.access_token)

    return redirect('/generate-certificates')


@app.route('/login/google')
def login_google():
    redirect_to = url_for('auth_callback_page', _external=True)
    result = supabase.auth.sign_in_with_oauth({
        "provider": "google",
        "options": {"redirect_to": redirect_to}
    })
    return redirect(result.url)


@app.route('/auth/callback-page')
def auth_callback_page():
    # Supabase's default OAuth flow (PKCE) redirects here with a `code`
    # query parameter, which the server CAN read directly -- no JS needed.
    code = request.args.get('code')

    if not code:
        return render_template('login.html', error="Google sign-in failed: no code received.")

    try:
        result = supabase.auth.exchange_code_for_session({"auth_code": code})
    except Exception:
        return render_template('login.html', error="Google sign-in failed. Please try again.")

    session['user_id'] = result.user.id
    session['email'] = result.user.email
    session['access_token'] = result.session.access_token

    update_last_sign_in(result.user.id, result.session.access_token)

    return redirect('/generate-certificates')


@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')


# ---------------------------------------------------------------------------
# App routes
# ---------------------------------------------------------------------------

@app.route('/generate-certificates')
@login_required
@subscription_required
def home():
    profile = get_profile_for_user(session['user_id'], session['access_token'])
    is_super_admin = profile.get('role') == 'super_admin' if profile else False
    is_admin = profile.get('role') in ('super_admin', 'sub_admin') if profile else False
    is_subscribed = profile.get('subscription_status') == 'active' if profile else False
    trial_remaining = FREE_TRIAL_LIMIT - (profile.get('free_certificates_used') or 0) if profile else 0

    return render_template(
        'index.html',
        is_super_admin=is_super_admin,
        is_admin=is_admin,
        is_subscribed=is_subscribed,
        trial_remaining=max(trial_remaining, 0)
    )

@app.route('/subscription-required')
@login_required
def subscription_required_page():
    return render_template('subscription_required.html')


@app.route('/generate', methods=['POST'])
@login_required
@subscription_required
def generate():
    template_image = request.files['template_image']
    excel_file = request.files['excel_file']
    x_percent = float(request.form['x_percent'])
    y_percent = float(request.form['y_percent'])
    font_file = request.form['font_file']
    font_size = int(float(request.form['font_size']))
    text_color = request.form['text_color']  # e.g. "#ff0000"

    # Save uploads
    image_path = os.path.join(UPLOAD_FOLDER, template_image.filename)
    excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
    template_image.save(image_path)
    excel_file.save(excel_path)

    # Read names from Excel
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active
    headers = [cell.value for cell in sheet[1]]
    name_col_index = headers.index('Name')

    names = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[name_col_index]:
            names.append(str(row[name_col_index]))

    # Enforce free trial limit for non-admin, non-subscribed users.
    # Admins and active subscribers generate without limit.
    profile = get_profile_for_user(session['user_id'], session['access_token'])
    is_admin = profile.get('role') in ('super_admin', 'sub_admin')
    is_subscribed = profile.get('subscription_status') == 'active'

    if not is_admin and not is_subscribed:
        used = profile.get('free_certificates_used') or 0
        remaining = FREE_TRIAL_LIMIT - used
        if len(names) > remaining:
            return jsonify({
                "error": (
                    f"Your free trial allows {remaining} more certificate(s), "
                    f"but this file has {len(names)} names. Subscribe to generate "
                    f"unlimited certificates, or upload a smaller list."
                )
            }), 403

    # Load font (fallback to Arial if the chosen font file isn't found)
    try:
        font = ImageFont.truetype(font_file, font_size)
    except Exception:
        font = ImageFont.truetype("arial.ttf", font_size)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for name in names:
            img = Image.open(image_path).convert("RGB")
            draw = ImageDraw.Draw(img)

            x = (x_percent / 100) * img.width
            y = (y_percent / 100) * img.height

            bbox = draw.textbbox((0, 0), name, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]

            # x,y is the CENTER (matches the draggable preview), shift to top-left for drawing
            draw_x = x - text_width / 2
            draw_y = y - text_height / 2

            draw.text((draw_x, draw_y), name, fill=text_color, font=font)

            pdf_buffer = io.BytesIO()
            img.save(pdf_buffer, format="PDF")
            pdf_buffer.seek(0)
            zip_file.writestr(f"Certificate_{name}.pdf", pdf_buffer.read())

    zip_buffer.seek(0)

    if not is_admin and not is_subscribed:
        increment_free_certificates_used(session['user_id'], session['access_token'], len(names))

    return send_file(
        zip_buffer,
        mimetype='application/zip',
        as_attachment=True,
        download_name='certificates.zip'
    )


# ---------------------------------------------------------------------------
# Admin panel routes
# ---------------------------------------------------------------------------

@app.route('/admin')
@login_required
@admin_required
def admin_panel():
    client = create_client(SUPABASE_URL, SUPABASE_KEY)
    client.postgrest.auth(session['access_token'])
    result = client.table('profiles').select('*').order('created_at', desc=True).execute()
    return render_template('admin.html', users=result.data, current_user_id=session['user_id'])


@app.route('/admin/toggle-sub-admin', methods=['POST'])
@login_required
@admin_required
def toggle_sub_admin():
    target_user_id = request.form.get('user_id')
    new_role = request.form.get('new_role')

    # Only allow toggling between 'user' and 'sub_admin'.
    # Promoting/demoting super_admin is intentionally NOT exposed here --
    # that stays a manual SQL action to prevent accidental loss of admin access.
    if new_role not in ('user', 'sub_admin'):
        return redirect('/admin')

    client = create_client(SUPABASE_URL, SUPABASE_KEY)
    client.postgrest.auth(session['access_token'])

    # Guard: never let this endpoint touch a super_admin row, even if someone
    # tampers with the form data client-side.
    target = client.table('profiles').select('role').eq('id', target_user_id).single().execute()
    if target.data and target.data.get('role') == 'super_admin':
        return redirect('/admin')

    client.table('profiles').update({'role': new_role}).eq('id', target_user_id).execute()
    return redirect('/admin')


if __name__ == '__main__':
    app.run(debug=True)


