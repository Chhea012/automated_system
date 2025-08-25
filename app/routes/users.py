import os
from flask import Blueprint, render_template, request, redirect, url_for, flash, current_app, jsonify
from flask_login import login_required
from werkzeug.utils import secure_filename
from app import db
from app.models.user import User
from app.models.role import Role
from app.models.department import Department

users_bp = Blueprint("users", __name__)

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@users_bp.route("/")
@login_required
def index():
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '', type=str)
    role_id = request.args.get('role_id', '', type=str)
    sort = request.args.get('sort', 'asc', type=str)

    # Base query
    query = User.query

    # Apply search filter
    if search:
        query = query.filter(User.username.ilike(f'%{search}%'))

    # Apply role filter
    if role_id:
        query = query.filter(User.role_id == role_id)

    # Apply sorting
    if sort == 'desc':
        query = query.order_by(User.username.desc())
    else:
        query = query.order_by(User.username.asc())

    # Paginate the results
    pagination = query.paginate(page=page, per_page=7, error_out=False)

    # Calculate metrics for the cards
    total_users = User.query.count()
    total_hr_admin = User.query.join(Role).filter(Role.name.in_(['HR', 'Admin'])).count()
    total_managers = User.query.join(Role).filter(Role.name == 'Manager').count()
    total_employees = User.query.join(Role).filter(Role.name == 'Employee').count()

    roles = Role.query.all()
    departments = Department.query.all()

    # Handle AJAX request
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return jsonify({
            'users': [
                {
                    'id': user.id,
                    'username': user.username or 'N/A',
                    'email': user.email or 'N/A',
                    'phone_number': user.phone_number or 'N/A',
                    'role_name': user.role.name if user.role else 'N/A',
                    'department_name': user.department.name if user.department else 'N/A',
                    'address': user.address or 'N/A',
                    'image_url': user.get_image_url()
                } for user in pagination.items
            ],
            'pagination': {
                'has_prev': pagination.has_prev,
                'has_next': pagination.has_next,
                'prev_num': pagination.prev_num,
                'next_num': pagination.next_num,
                'page': pagination.page,
                'pages': pagination.pages,
                'iter_pages': list(pagination.iter_pages())
            }
        })

    return render_template(
        "users/index.html",
        users=pagination.items,
        pagination=pagination,
        roles=roles,
        departments=departments,
        total_users=total_users,
        total_hr_admin=total_hr_admin,
        total_managers=total_managers,
        total_employees=total_employees,
        search=search,
        role_id=role_id,
        sort=sort
    )

@users_bp.route("/profile/<int:user_id>")
@login_required
def profile(user_id):
    user = User.query.get_or_404(user_id)
    roles = Role.query.all()
    departments = Department.query.all()
    return render_template(
        "users/profile.html",
        user=user,
        roles=roles,
        departments=departments,
        page=request.args.get('page', 1, type=int)
    )

@users_bp.route("/create", methods=["POST"])
@login_required
def create():
    username = request.form.get("username")
    email = request.form.get("email")
    password = request.form.get("password")
    phone_number = request.form.get("phone_number")
    address = request.form.get("address")
    role_id = request.form.get("role_id")
    department_id = request.form.get("department_id")
    image = request.files.get("image")

    if not username or not email or not password:
        flash("Username, email, and password are required!", "danger")
        return redirect(url_for("users.index", page=1))

    if User.query.filter_by(username=username).first():
        flash("Username already exists!", "danger")
        return redirect(url_for("users.index", page=1))

    if User.query.filter_by(email=email).first():
        flash("Email already exists!", "danger")
        return redirect(url_for("users.index", page=1))

    new_user = User(username=username, email=email)
    new_user.set_password(password)
    new_user.phone_number = phone_number
    new_user.address = address
    new_user.role_id = int(role_id) if role_id else None
    new_user.department_id = int(department_id) if department_id else None

    if image and allowed_file(image.filename):
        filename = secure_filename(image.filename)
        upload_folder = os.path.join(current_app.root_path, 'static/uploads')
        os.makedirs(upload_folder, exist_ok=True)
        file_path = os.path.join(upload_folder, filename)
        image.save(file_path)
        new_user.image = filename

    db.session.add(new_user)
    db.session.commit()
    flash("User created successfully!", "success")
    return redirect(url_for("users.index", page=1))

@users_bp.route("/update/<int:user_id>", methods=["POST"])
@login_required
def update(user_id):
    user = User.query.get_or_404(user_id)
    username = request.form.get("username")
    email = request.form.get("email")
    password = request.form.get("password")
    phone_number = request.form.get("phone_number")
    address = request.form.get("address")
    role_id = request.form.get("role_id")
    department_id = request.form.get("department_id")
    image = request.files.get("image")

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
    user.role_id = int(role_id) if role_id else None
    user.department_id = int(department_id) if department_id else None

    if password:
        user.set_password(password)

    if image and allowed_file(image.filename):
        filename = secure_filename(image.filename)
        upload_folder = os.path.join(current_app.root_path, 'static/uploads')
        os.makedirs(upload_folder, exist_ok=True)
        file_path = os.path.join(upload_folder, filename)
        image.save(file_path)
        user.image = filename

    db.session.commit()
    flash("User updated successfully!", "success")
    return redirect(url_for("users.index", page=request.args.get('page', 1)))

@users_bp.route("/delete/<int:user_id>", methods=["POST"])
@login_required
def delete(user_id):
    user = User.query.get_or_404(user_id)
    db.session.delete(user)
    db.session.commit()
    flash("User deleted!", "success")
    return redirect(url_for("users.index", page=request.args.get('page', 1)))