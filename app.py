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
import smtplib
from email.mime.text import MIMEText

# Load environment variables from .env file
load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
FLASK_SECRET_KEY = os.getenv("FLASK_SECRET_KEY")
INTRO_VIDEO_BUCKET = 'intro-video'
ALLOWED_VIDEO_EXTENSIONS = {'.mp4', '.webm'}
MAX_VIDEO_SIZE_BYTES = 200 * 1024 * 1024  # 200MB
SIGNED_URL_EXPIRY_SECONDS = 60 * 60 * 24  # 24 hours

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
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

FREE_TRIAL_LIMIT = 5  # Number of free certificates allowed for trial users

LEMONSQUEEZY_API_KEY = os.getenv("LEMONSQUEEZY_API_KEY")
LEMONSQUEEZY_STORE_ID = os.getenv("LEMONSQUEEZY_STORE_ID")
LEMONSQUEEZY_WEBHOOK_SECRET = os.getenv("LEMONSQUEEZY_WEBHOOK_SECRET")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY")

VARIANT_IDS = {
    'monthly': os.getenv("LEMONSQUEEZY_VARIANT_MONTHLY"),
    'sixmonth': os.getenv("LEMONSQUEEZY_VARIANT_6MONTH"),
    'yearly': os.getenv("LEMONSQUEEZY_VARIANT_YEARLY"),
}

SMTP_EMAIL = os.getenv("SMTP_EMAIL")
SMTP_APP_PASSWORD = os.getenv("SMTP_APP_PASSWORD")
NOTIFY_EMAIL = os.getenv("NOTIFY_EMAIL")

PLAN_DETAILS = {
    'monthly': {'name': 'Monthly', 'price': '9.99'},
    'sixmonth': {'name': '6 Months', 'price': '49.99'},
    'yearly': {'name': 'Yearly', 'price': '84.99'},
}

def send_subscription_request_email(full_name, phone, email, plan_name, plan_price):
    subject = f"New Subscription Request - {plan_name} (${plan_price})"
    body = (
        f"New subscription request received:\n\n"
        f"Name: {full_name}\n"
        f"Phone: {phone}\n"
        f"Email: {email}\n"
        f"Plan: {plan_name} - ${plan_price}\n\n"
        f"Once payment is confirmed via JazzCash, grant access from /admin "
        f"by searching this email and clicking 'Grant Free Access'."
    )
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = SMTP_EMAIL
    msg['To'] = NOTIFY_EMAIL

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(SMTP_EMAIL, SMTP_APP_PASSWORD)
        server.sendmail(SMTP_EMAIL, NOTIFY_EMAIL, msg.as_string())
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

def extract_names_from_excel(excel_path):
    """
    Forgiving Excel reader: finds a 'Name' column regardless of case or
    surrounding spaces, skips fully blank leading rows, and falls back to
    reading the first column directly if no 'Name' header is found at all
    (covers files with no header row, or a header the user didn't label).
    """
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active

    rows = list(sheet.iter_rows(values_only=True))
    # Skip any fully blank rows at the top (e.g. an accidental empty first row)
    rows = [row for row in rows if any(cell is not None and str(cell).strip() != '' for cell in row)]

    if not rows:
        return []

    header_row = rows[0]
    name_col_index = None
    for i, cell in enumerate(header_row):
        if cell and str(cell).strip().lower() == 'name':
            name_col_index = i
            break

    if name_col_index is not None:
        # Real header found -- skip it, read the rest as data
        data_rows = rows[1:]
    else:
        # No 'Name' header found anywhere -- assume no header row exists,
        # and that the first column holds the names directly.
        name_col_index = 0
        data_rows = rows

    names = []
    for row in data_rows:
        if len(row) > name_col_index and row[name_col_index] is not None:
            value = str(row[name_col_index]).strip()
            if value:
                names.append(value)

    return names

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

@app.route('/health')
def health_check():
    return jsonify({'status': 'ok'}), 200

'''
# LEMON SQUEEZY -- DISABLED (switched to manual JazzCash flow below).
# Kept here, commented out, in case live-mode payment gateway issues get
# resolved later and we want to re-enable automatic checkout.

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

    return jsonify({'received': True}), 200  '''

def admin_required(view_func):
    @functools.wraps(view_func)
    def wrapped(*args, **kwargs):
        profile = get_profile_for_user(session['user_id'], session['access_token'])
        if not profile or profile.get('role') != 'super_admin':
            return redirect('/generate-certificates')
        return view_func(*args, **kwargs)
    return wrapped

def get_intro_video_info():
    client = get_service_client()
    for ext in ALLOWED_VIDEO_EXTENSIONS:
        object_path = f'intro{ext}'
        try:
            files = client.storage.from_(INTRO_VIDEO_BUCKET).list()
            if any(f['name'] == object_path for f in files):
                result = client.storage.from_(INTRO_VIDEO_BUCKET).create_signed_url(
                    object_path, SIGNED_URL_EXPIRY_SECONDS
                )
                return result['signedURL'], ext.lstrip('.')
        except Exception:
            continue
    return None, None

# ---------------------------------------------------------------------------
# Public routes
# ---------------------------------------------------------------------------

@app.route('/')
def landing():
    is_logged_in = bool(session.get('user_id'))
    intro_video_url, intro_video_ext = get_intro_video_info()
    return render_template(
        'landing.html',
        is_logged_in=is_logged_in,
        intro_video_url=intro_video_url,
        intro_video_ext=intro_video_ext
    )


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


@app.route('/subscribe-request')
@login_required
def subscribe_request():
    plan = request.args.get('plan')
    if plan not in PLAN_DETAILS:
        return redirect('/subscription-required')
    return render_template(
        'subscribe_request.html',
        plan=plan,
        plan_name=PLAN_DETAILS[plan]['name'],
        plan_price=PLAN_DETAILS[plan]['price']
    )


@app.route('/submit-subscription-request', methods=['POST'])
@login_required
def submit_subscription_request():
    plan = request.form.get('plan')
    full_name = request.form.get('full_name', '').strip()
    phone = request.form.get('phone', '').strip()
    email = request.form.get('email', '').strip()

    if plan not in PLAN_DETAILS or not full_name or not phone or not email:
        return redirect('/subscription-required')

    plan_info = PLAN_DETAILS[plan]
    try:
        send_subscription_request_email(full_name, phone, email, plan_info['name'], plan_info['price'])
    except Exception as e:
        app.logger.error(f"Failed to send subscription request email: {e}")

    return render_template('subscription_processing.html')


@app.route('/generate', methods=['POST'])
@login_required
@subscription_required
def generate():
    template_image = request.files['template_image']
    input_mode = request.form.get('input_mode', 'excel')
    x_percent = float(request.form['x_percent'])
    y_percent = float(request.form['y_percent'])
    font_file = request.form['font_file']
    font_size = int(float(request.form['font_size']))
    text_color = request.form['text_color']  # e.g. "#ff0000"

    # Save the template image (always required regardless of input mode)
    image_path = os.path.join(UPLOAD_FOLDER, template_image.filename)
    template_image.save(image_path)

    if input_mode == 'paste':
        pasted_text = request.form.get('pasted_names', '')
        names = [line.strip() for line in pasted_text.splitlines() if line.strip()]
    else:
        excel_file = request.files['excel_file']
        excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
        excel_file.save(excel_path)
        names = extract_names_from_excel(excel_path)

    if not names:
        return jsonify({"error": "No names found. Please check your file or pasted list."}), 400

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
        font = ImageFont.truetype("static/fonts/Arimo-Regular.ttf", font_size)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for name in names:
            img = Image.open(image_path).convert("RGB")
            draw = ImageDraw.Draw(img)

            x = (x_percent / 100) * img.width
            y = (y_percent / 100) * img.height

            draw.text((x, y), name, fill=text_color, font=font, anchor="mm")

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

ALLOWED_VIDEO_EXTENSIONS = {'.mp4', '.webm'}
MAX_VIDEO_SIZE_BYTES = 200 * 1024 * 1024  # 200MB


@app.route('/admin/upload-video', methods=['POST'])
@login_required
@admin_required
def upload_video():
    video_file = request.files.get('intro_video')
    if not video_file or video_file.filename == '':
        return redirect('/admin')

    ext = os.path.splitext(video_file.filename)[1].lower()
    if ext not in ALLOWED_VIDEO_EXTENSIONS:
        return "Invalid file type. Please upload an .mp4 or .webm file.", 400

    video_file.seek(0, os.SEEK_END)
    file_size = video_file.tell()
    video_file.seek(0)
    if file_size > MAX_VIDEO_SIZE_BYTES:
        return "File too large. Maximum size is 200MB.", 400

    client = get_service_client()

    for other_ext in ALLOWED_VIDEO_EXTENSIONS:
        if other_ext != ext:
            try:
                client.storage.from_(INTRO_VIDEO_BUCKET).remove([f'intro{other_ext}'])
            except Exception:
                pass

    file_bytes = video_file.read()
    object_path = f'intro{ext}'
    content_type = 'video/mp4' if ext == '.mp4' else 'video/webm'

    client.storage.from_(INTRO_VIDEO_BUCKET).upload(
        object_path,
        file_bytes,
        file_options={"content-type": content_type, "upsert": "true"}
    )

    return redirect('/admin')

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


