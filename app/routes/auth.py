from flask import Blueprint, render_template, redirect, url_for, flash
from flask_login import login_user, logout_user, login_required
from .. import db
from ..models.user import User
from ..forms import LoginForm, RegisterForm

auth_bp = Blueprint("auth", __name__)

@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    form = LoginForm()
    if form.validate_on_submit():
        identifier = form.identifier.data
        print(f"Login attempt: identifier={identifier}")  # Debug
        user = User.query.filter((User.email == identifier) | (User.username == identifier)).first()
        if user:
            print(f"User found: {user.username}, hash: {user.password_hash}")  # Debug
            if user.check_password(form.password.data):
                print("Password check: SUCCESS")  # Debug
                login_user(user, remember=form.remember.data)
                flash("Login successful! Redirecting to dashboard.", "success")
                return redirect(url_for("main.dashboard"))
            else:
                print("Password check: FAILED")  # Debug
                flash("Invalid identifier or password. Please try again.", "danger")
        else:
            print("No user found for identifier")  # Debug
            flash("Invalid identifier or password. Please try again.", "danger")
    return render_template("auth/login.html", form=form)

@auth_bp.route("/register", methods=["GET", "POST"])
def register():
    form = RegisterForm()
    if form.validate_on_submit():
        username = form.username.data
        email = form.email.data
        print(f"Registering: username={username}, email={email}")  # Debug
        if User.query.filter_by(email=email).first():
            flash("Email already registered. Please use a different email.", "warning")
            return redirect(url_for("auth.register"))
        new_user = User(username=username, email=email)
        new_user.set_password(form.password.data)
        try:
            db.session.add(new_user)
            db.session.commit()
            print(f"User {username} saved with hash: {new_user.password_hash}")  # Debug
            login_user(new_user)
            flash("Registration successful! Redirecting to dashboard.", "success")
            return redirect(url_for("main.dashboard"))
        except Exception as e:
            db.session.rollback()
            print(f"Registration error: {str(e)}")  # Debug
            flash("Registration failed. Please try again later.", "danger")
    return render_template("auth/register.html", form=form)

@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    flash("You have been logged out.", "info")
    return redirect(url_for("auth.login"))