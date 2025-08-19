from flask import Blueprint, render_template, redirect, url_for, request, flash
from flask_login import login_user, logout_user, login_required
from .. import db
from ..models.user import User

auth_bp = Blueprint("auth", __name__)

@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        identifier = request.form.get("identifier")
        password = request.form.get("password")
        user = User.query.filter((User.email == identifier) | (User.username == identifier)).first()
        if user and user.check_password(password):
            login_user(user, remember=request.form.get("remember") == "on")
            flash("Login successful! Redirecting to dashboard.", "success")
            return redirect(url_for("main.dashboard"))
        else:
            flash("Invalid identifier or password. Please try again.", "danger")
    return render_template("auth/login.html")

@auth_bp.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form.get("username")
        email = request.form.get("email")
        password = request.form.get("password")
        if User.query.filter_by(email=email).first():
            flash("Email already registered. Please use a different email.", "warning")
            return redirect(url_for("auth.register"))
        new_user = User(username=username, email=email)
        new_user.set_password(password)
        try:
            db.session.add(new_user)
            db.session.commit()
            login_user(new_user)
            flash("Registration successful! Redirecting to dashboard.", "success")
            return redirect(url_for("main.dashboard"))
        except Exception as e:
            db.session.rollback()
            flash("Registration failed. Please try again later.", "danger")
    return render_template("auth/register.html")

@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    flash("You have been logged out.", "info")
    return redirect(url_for("auth.login"))