from flask import Blueprint, render_template, redirect, url_for, flash, request, current_app
from flask_login import login_user, logout_user, login_required, current_user
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from itsdangerous import URLSafeTimedSerializer, SignatureExpired, BadSignature
from flask_mail import Message
from .. import db, mail
from ..models.user import User
from ..forms import LoginForm, RegisterForm, PasswordResetRequestForm, PasswordResetForm

auth_bp = Blueprint("auth", __name__)
limiter = Limiter(key_func=get_remote_address)

# ------------------------------
# Token Helpers
# ------------------------------
def generate_reset_token(email):
    serializer = URLSafeTimedSerializer(current_app.config['SECRET_KEY'])
    return serializer.dumps(email, salt='password-reset-salt')

def verify_reset_token(token, expiration=3600):
    serializer = URLSafeTimedSerializer(current_app.config['SECRET_KEY'])
    try:
        email = serializer.loads(token, salt='password-reset-salt', max_age=expiration)
        return email
    except (SignatureExpired, BadSignature):
        return None

# ------------------------------
# Login
# ------------------------------
@auth_bp.route("/login", methods=["GET", "POST"])
@limiter.limit("5 per minute")
def login():
    if current_user.is_authenticated:
        return redirect(url_for("main.dashboard"))

    form = LoginForm()
    if form.validate_on_submit():
        identifier = form.identifier.data
        user = User.query.filter((User.email == identifier) | (User.username == identifier)).first()

        if user and user.check_password(form.password.data):
            login_user(user, remember=form.remember.data)
            flash("Login successful! Redirecting to dashboard.", "success")
            return redirect(url_for("main.dashboard"))

        flash("Invalid identifier or password. Please try again.", "danger")

    return render_template("auth/login.html", form=form)

# ------------------------------
# Register
# ------------------------------
@auth_bp.route("/register", methods=["GET", "POST"])
def register():
    if current_user.is_authenticated:
        return redirect(url_for("main.dashboard"))

    form = RegisterForm()
    if form.validate_on_submit():
        username = form.username.data
        email = form.email.data

        if User.query.filter_by(email=email).first():
            flash("Email already registered. Please use a different email.", "warning")
            return redirect(url_for("auth.register"))

        try:
            new_user = User(username=username, email=email)
            new_user.set_password(form.password.data)
            db.session.add(new_user)
            db.session.commit()
            login_user(new_user)
            flash("Registration successful! Redirecting to dashboard.", "success")
            return redirect(url_for("main.dashboard"))
        except Exception as e:
            db.session.rollback()
            current_app.logger.error(f"Registration error: {str(e)}")
            flash("Registration failed. Please try again later.", "danger")

    return render_template("auth/register.html", form=form)

# ------------------------------
# Logout
# ------------------------------
@auth_bp.route("/logout")
@login_required
def logout():
    try:
        logout_user()
        flash("You have been logged out.", "info")
        return redirect(url_for("auth.login"))
    except Exception as e:
        current_app.logger.error(f"Logout error: {str(e)}")
        flash("An error occurred during logout. Please try again.", "danger")
        return redirect(url_for("main.dashboard"))

# ------------------------------
# Reset Password Request
# ------------------------------
@auth_bp.route("/reset_password_request", methods=["GET", "POST"])
@limiter.limit("5 per hour")
def reset_password_request():
    if current_user.is_authenticated:
        return redirect(url_for("main.dashboard"))

    form = PasswordResetRequestForm()
    if form.validate_on_submit():
        email = form.email.data
        user = User.query.filter_by(email=email).first()
        if user:
            token = generate_reset_token(email)
            reset_url = url_for('auth.reset_password', token=token, _external=True)
            msg = Message('Password Reset Request', recipients=[email])
            msg.body = f'''To reset your password, click the following link: {reset_url}
If you did not request this, please ignore this email.
The link will expire in 1 hour.'''
            mail.send(msg)
            flash("A password reset link has been sent to your email.", "success")
        else:
            flash("Email not found. Please check and try again.", "warning")

        return redirect(url_for("auth.login"))

    return render_template("auth/reset_password_request.html", form=form)

# ------------------------------
# Reset Password (with token)
# ------------------------------
@auth_bp.route("/reset_password/<token>", methods=["GET", "POST"])
@limiter.limit("5 per hour")
def reset_password(token):
    if current_user.is_authenticated:
        return redirect(url_for("main.dashboard"))

    email = verify_reset_token(token)
    if not email:
        flash("The reset link is invalid or has expired.", "danger")
        return redirect(url_for("auth.login"))

    user = User.query.filter_by(email=email).first()
    if not user:
        flash("User not found.", "danger")
        return redirect(url_for("auth.login"))

    form = PasswordResetForm()
    if form.validate_on_submit():
        try:
            user.set_password(form.password.data)
            db.session.commit()
            login_user(user)
            flash("Your password has been reset and you are now logged in.", "success")
            return redirect(url_for("main.dashboard"))
        except Exception as e:
            db.session.rollback()
            current_app.logger.error(f"Password reset error: {str(e)}")
            flash("An error occurred while resetting your password.", "danger")

    return render_template("auth/reset_password.html", form=form, token=token)
