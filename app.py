from flask import Flask, render_template, request, jsonify, redirect, url_for, session, flash
import sqlite3
import hashlib
import random
import time
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'dev-secret-key')

MAIL_USERNAME = os.getenv("MAIL_USERNAME")
MAIL_PASSWORD = os.getenv("MAIL_PASSWORD")

DB_PATH = "database/chatbot.db"

# ════════════════════════════════════════════════
#  DB HELPERS
# ════════════════════════════════════════════════

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def hash_password(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

# ════════════════════════════════════════════════
#  OTP HELPERS
# ════════════════════════════════════════════════

def generate_otp():
    return str(random.randint(100000, 999999))

def send_email_otp(to_email, otp):
    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = 'StudentAI — Your OTP Code'
        msg['From']    = MAIL_USERNAME
        msg['To']      = to_email

        html = f"""
        <div style="font-family:Arial,sans-serif;max-width:480px;margin:auto;background:#0a0a0f;color:#f0f0f0;border-radius:16px;padding:40px;border:1px solid rgba(255,255,255,0.1);">
          <h2 style="font-family:Georgia,serif;color:#4da6ff;margin-bottom:8px;">StudentAI</h2>
          <p style="color:rgba(255,255,255,0.5);margin-bottom:32px;font-size:14px;">Your verification code</p>
          <div style="background:#16161f;border-radius:12px;padding:24px;text-align:center;border:1px solid rgba(77,166,255,0.2);margin-bottom:24px;">
            <span style="font-size:42px;font-weight:900;letter-spacing:12px;color:#4da6ff;">{otp}</span>
          </div>
          <p style="color:rgba(255,255,255,0.4);font-size:13px;">This code expires in <strong style="color:#f0f0f0;">10 minutes</strong>. Do not share it with anyone.</p>
        </div>
        """
        msg.attach(MIMEText(html, 'html'))

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(MAIL_USERNAME, MAIL_PASSWORD)
            server.sendmail(MAIL_USERNAME, to_email, msg.as_string())
        print(f"[Email] OTP sent to {to_email}")
        return True
    except Exception as e:
        print(f"[Email] Error: {e}")
        return False

# ════════════════════════════════════════════════
#  CHATBOT
# ════════════════════════════════════════════════

try:
    from chatbot.model import chatbot_response
except ImportError:
    def chatbot_response(msg, mode="student"):
        return f"Echo: {msg}  [chatbot model not loaded]"

# ════════════════════════════════════════════════
#  ROUTES
# ════════════════════════════════════════════════

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/app')
def index():
    if 'username' not in session:
        return redirect(url_for('login'))
    return render_template('index.html', username=session['username'])

# ── SIGNUP ──
@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        email    = request.form.get('email', '').strip()
        password = request.form.get('password', '').strip()
        confirm  = request.form.get('confirm_password', '').strip()

        if not all([username, email, password, confirm]):
            flash('Please fill in all fields.', 'error')
            return render_template('signup.html')
        if password != confirm:
            flash('Passwords do not match.', 'error')
            return render_template('signup.html')

        db = get_db()
        existing = db.execute(
            "SELECT id FROM users WHERE username=? OR email=?",
            (username, email)
        ).fetchone()
        db.close()
        if existing:
            flash('Username or email already taken.', 'error')
            return render_template('signup.html')

        otp = generate_otp()
        session['signup_otp']      = otp
        session['signup_otp_time'] = time.time()
        session['signup_data']     = {
            'username': username,
            'email':    email,
            'password': hash_password(password)
        }

        if not send_email_otp(email, otp):
            flash('Could not send OTP. Check your email and try again.', 'error')
            return render_template('signup.html')

        flash('OTP sent to your email. Enter it below.', 'success')
        return redirect(url_for('verify_signup_otp'))

    return render_template('signup.html')


@app.route('/verify-signup', methods=['GET', 'POST'])
def verify_signup_otp():
    if 'signup_otp' not in session:
        return redirect(url_for('signup'))

    if request.method == 'POST':
        entered = request.form.get('otp', '').strip()
        stored  = session.get('signup_otp')
        sent_at = session.get('signup_otp_time', 0)

        if time.time() - sent_at > 600:
            session.pop('signup_otp', None)
            flash('OTP expired. Please sign up again.', 'error')
            return redirect(url_for('signup'))

        if entered != stored:
            flash('Incorrect OTP. Please try again.', 'error')
            return render_template('verify_otp.html', action='signup')

        data = session.pop('signup_data')
        session.pop('signup_otp', None)
        session.pop('signup_otp_time', None)

        try:
            db = get_db()
            db.execute(
                "INSERT INTO users (username, email, password) VALUES (?,?,?)",
                (data['username'], data['email'], data['password'])
            )
            db.commit()
            db.close()
            flash('Account created! Please sign in.', 'success')
            return redirect(url_for('login'))
        except sqlite3.IntegrityError:
            flash('Username or email already taken.', 'error')
            return redirect(url_for('signup'))

    return render_template('verify_otp.html', action='signup')

# ── LOGIN ──
@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'username' in session:
        return redirect(url_for('index'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        if not username or not password:
            flash('Please fill in all fields.', 'error')
            return render_template('login.html')
        db = get_db()
        user = db.execute(
            "SELECT * FROM users WHERE username=? AND password=?",
            (username, hash_password(password))
        ).fetchone()
        db.close()
        if user:
            session['username'] = username
            return redirect(url_for('index'))
        else:
            flash('Invalid username or password.', 'error')
    return render_template('login.html')

# ── FORGOT PASSWORD ──
@app.route('/forgot-password', methods=['GET', 'POST'])
def forgot_password():
    if request.method == 'POST':
        identifier = request.form.get('identifier', '').strip()

        db = get_db()
        user = db.execute(
            "SELECT * FROM users WHERE email=? OR username=?",
            (identifier, identifier)
        ).fetchone()
        db.close()

        if not user:
            flash('No account found with that email or username.', 'error')
            return render_template('forgot_password.html')

        otp = generate_otp()
        session['reset_otp']      = otp
        session['reset_otp_time'] = time.time()
        session['reset_username'] = user['username']
        session['reset_email']    = user['email']

        if not send_email_otp(user['email'], otp):
            flash('Could not send OTP. Try again.', 'error')
            return render_template('forgot_password.html')

        flash('OTP sent to your registered email.', 'success')
        return redirect(url_for('verify_reset_otp'))

    return render_template('forgot_password.html')


@app.route('/verify-reset', methods=['GET', 'POST'])
def verify_reset_otp():
    if 'reset_otp' not in session:
        return redirect(url_for('forgot_password'))

    if request.method == 'POST':
        entered = request.form.get('otp', '').strip()
        stored  = session.get('reset_otp')
        sent_at = session.get('reset_otp_time', 0)

        if time.time() - sent_at > 600:
            session.pop('reset_otp', None)
            flash('OTP expired. Please try again.', 'error')
            return redirect(url_for('forgot_password'))

        if entered != stored:
            flash('Incorrect OTP.', 'error')
            return render_template('verify_otp.html', action='reset')

        session.pop('reset_otp', None)
        session.pop('reset_otp_time', None)
        session['reset_verified'] = True
        return redirect(url_for('reset_password'))

    return render_template('verify_otp.html', action='reset')


@app.route('/reset-password', methods=['GET', 'POST'])
def reset_password():
    if not session.get('reset_verified'):
        return redirect(url_for('forgot_password'))

    if request.method == 'POST':
        password = request.form.get('password', '').strip()
        confirm  = request.form.get('confirm_password', '').strip()

        if not password or not confirm:
            flash('Please fill in both fields.', 'error')
            return render_template('reset_password.html')
        if password != confirm:
            flash('Passwords do not match.', 'error')
            return render_template('reset_password.html')

        username = session.pop('reset_username')
        session.pop('reset_verified', None)
        session.pop('reset_email', None)

        db = get_db()
        db.execute(
            "UPDATE users SET password=? WHERE username=?",
            (hash_password(password), username)
        )
        db.commit()
        db.close()

        flash('Password reset successfully! Please sign in.', 'success')
        return redirect(url_for('login'))

    return render_template('reset_password.html')

# ── LOGOUT + CHAT ──
@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('home'))

@app.route('/chat', methods=['POST'])
def chat():
    if 'username' not in session:
        return jsonify({"error": "Unauthorized"}), 401
    user_message = request.json.get('message', '')
    mode = request.json.get('mode', 'student')
    reply = chatbot_response(user_message, mode)
    try:
        db = get_db()
        db.execute("INSERT INTO messages (sender, message) VALUES (?,?)", ('user', user_message))
        db.execute("INSERT INTO messages (sender, message) VALUES (?,?)", ('bot', reply))
        db.commit()
        db.close()
    except Exception:
        pass
    return jsonify({"reply": reply})

@app.route('/about')
def about():
    return render_template('about.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)