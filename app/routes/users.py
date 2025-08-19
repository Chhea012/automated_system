from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required
from app import db
from app.models.user import User
from app.utils.file_upload import save_profile_image

users_bp = Blueprint("users", __name__)

# List users with pagination, sorted by created_at descending
@users_bp.route("/")
@login_required
def index():
    page = request.args.get('page', 1, type=int)
    pagination = User.query.order_by(User.created_at.desc()).paginate(page=page, per_page=5, error_out=False)
    return render_template("users/index.html", users=pagination.items, pagination=pagination)

# Create user
@users_bp.route("/create", methods=["POST"])
@login_required
def create():
    username = request.form.get("username")
    email = request.form.get("email")
    password = request.form.get("password")
    phone_number = request.form.get("phone_number")
    address = request.form.get("address")
    file = request.files.get("profile_image")

    if not username or not email or not password:
        flash("Username, email, and password are required!", "danger")
        return redirect(url_for("users.index", page=1))

    if User.query.filter_by(username=username).first():
        flash("Username already exists!", "danger")
        return redirect(url_for("users.index", page=1))

    if User.query.filter_by(email=email).first():
        flash("Email already exists!", "danger")
        return redirect(url_for("users.index", page=1))

    new_user = User(
        username=username,
        email=email,
        phone_number=phone_number,
        address=address
    )
    new_user.set_password(password)

    if file and file.filename:
        filename = save_profile_image(file, username)
        if filename:
            new_user.image = filename

    db.session.add(new_user)
    db.session.commit()
    flash("User created successfully!", "success")
    return redirect(url_for("users.index", page=1))

# Update user
@users_bp.route("/update/<int:user_id>", methods=["POST"])
@login_required
def update(user_id):
    user = User.query.get_or_404(user_id)
    username = request.form.get("username")
    email = request.form.get("email")
    phone_number = request.form.get("phone_number")
    address = request.form.get("address")
    file = request.files.get("profile_image")

    if not username or not email:
        flash("Username and email are required!", "danger")
        return redirect(url_for("users.index", page=request.args.get('page', 1)))

    if username != user.username and User.query.filter_by(username=username).first():
        flash("Username already exists!", "danger")
        return redirect(url_for("users.index", page=request.args.get('page', 1)))

    if email != user.email and User.query.filter_by(email=email).first():
        flash("Email already exists!", "danger")
        return redirect(url_for("users.index", page=request.args.get('page', 1)))

    user.username = username
    user.email = email
    user.phone_number = phone_number
    user.address = address

    if file and file.filename:
        filename = save_profile_image(file, username)
        if filename:
            user.image = filename

    db.session.commit()
    flash("User updated successfully!", "success")
    return redirect(url_for("users.index", page=request.args.get('page', 1)))

# Delete user
@users_bp.route("/delete/<int:user_id>", methods=["POST"])
@login_required
def delete(user_id):
    user = User.query.get_or_404(user_id)
    db.session.delete(user)
    db.session.commit()
    flash("User deleted!", "success")
    return redirect(url_for("users.index", page=request.args.get('page', 1)))