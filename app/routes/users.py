import os
from flask import Blueprint, render_template, request, redirect, url_for, flash, current_app, jsonify
from flask_login import login_required, current_user
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
    if not current_user.has_role('Admin') and not current_user.has_role('Manager'):
        flash("You do not have permission to access this page.", "danger")
        return redirect(url_for("main.dashboard"))

    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '', type=str).strip()
    role_id = request.args.get('role_id', '', type=str)
    sort = request.args.get('sort', 'username_asc', type=str)
    per_page = request.args.get('per_page', 7, type=int)

    allowed_per_page = [7, 10, 25, 50]
    if per_page not in allowed_per_page:
        per_page = 7

    query = User.query

    if search:
        query = query.filter(
            (User.username.ilike(f'%{search}%')) | 
            (User.email.ilike(f'%{search}%'))
        )

    if role_id:
        query = query.filter(User.role_id == role_id)

    if sort == 'username_desc':
        query = query.order_by(User.username.desc())
    elif sort == 'email_asc':
        query = query.order_by(User.email.asc())
    elif sort == 'email_desc':
        query = query.order_by(User.email.desc())
    else:
        query = query.order_by(User.username.asc())

    try:
        pagination = query.paginate(page=page, per_page=per_page, error_out=False)
    except:
        page = 1
        pagination = query.paginate(page=page, per_page=per_page, error_out=False)

    total_users = User.query.count()
    total_admins = User.query.join(Role).filter(Role.name == 'Admin').count()
    total_hr_admin = User.query.join(Role).filter(Role.name.in_(['HR', 'Admin'])).count()
    total_managers = User.query.join(Role).filter(Role.name == 'Manager').count()
    total_employees = User.query.join(Role).filter(Role.name == 'Employee').count()

    roles = Role.query.all()
    departments = Department.query.all()

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
                'per_page': pagination.per_page,
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
        total_admins=total_admins,
        total_hr_admin=total_hr_admin,
        total_managers=total_managers,
        total_employees=total_employees,
        search=search,
        role_id=role_id,
        sort=sort,
        per_page=per_page
    )

@users_bp.route("/<int:user_id>/json")
@login_required
def get_user(user_id):
    if not current_user.has_role('Admin') and not current_user.has_role('Manager'):
        return jsonify({'error': 'Unauthorized'}), 403
    user = User.query.get_or_404(user_id)
    return jsonify({
        'id': user.id,
        'username': user.username,
        'email': user.email,
        'phone_number': user.phone_number,
        'address': user.address,
        'role_id': user.role_id,
        'role_name': user.role.name if user.role else 'N/A',
        'department_id': user.department_id,
        'image_url': user.get_image_url()
    })

@users_bp.route("/check_phone", methods=["GET"])
@login_required
def check_phone():
    if not (current_user.has_role('Admin') or current_user.has_role('Manager')):
        return jsonify({'error': 'Unauthorized'}), 403

    phone_number = request.args.get("phone_number", "").strip()
    user_id = request.args.get("user_id", type=int)

    if not phone_number:
        return jsonify({'exists': False})

    query = User.query.filter_by(phone_number=phone_number)
    if user_id:
        query = query.filter(User.id != user_id)

    exists = query.first() is not None
    return jsonify({'exists': exists})

@users_bp.route("/profile/<int:user_id>")
@login_required
def profile(user_id):
    user = User.query.get_or_404(user_id)
    if not (current_user.id == user_id or current_user.has_role('Admin') or current_user.has_role('Manager')):
        flash("You do not have permission to view this profile.", "danger")
        return redirect(url_for("main.dashboard"))

    roles = Role.query.all()
    departments = Department.query.all()
    total_admins = User.query.join(Role).filter(Role.name == 'Admin').count()
    return render_template(
        "users/profile.html",
        user=user,
        roles=roles,
        departments=departments,
        total_admins=total_admins,
        page=request.args.get('page', 1, type=int)
    )

@users_bp.route("/create", methods=["POST"])
@login_required
def create():
    if not current_user.has_role('Admin'):
        flash("You do not have permission to create users.", "danger")
        return redirect(url_for("main.dashboard"))

    username = request.form.get("username")
    email = request.form.get("email")
    password = request.form.get("password")
    phone_number = request.form.get("phone_number").strip() if request.form.get("phone_number") else None
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

    if phone_number and User.query.filter_by(phone_number=phone_number).first():
        flash("Phone number already exists!", "danger")
        return redirect(url_for("users.index", page=1))

    if role_id:
        role = Role.query.get(role_id)
        if role and role.name == 'Admin' and User.query.join(Role).filter(Role.name == 'Admin').count() > 0:
            flash("Only one Admin is allowed in the system!", "danger")
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
        try:
            image.save(file_path)
            new_user.image = filename
        except Exception as e:
            flash(f"Error uploading image: {str(e)}", "danger")
            return redirect(url_for("users.index", page=1))

    try:
        db.session.add(new_user)
        db.session.commit()
        flash("User created successfully!", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error creating user: {str(e)}", "danger")

    return redirect(url_for("users.index", page=1))

@users_bp.route("/update/<int:user_id>", methods=["POST"])
@login_required
def update(user_id):
    if not current_user.has_role('Admin'):
        flash("You do not have permission to update users.", "danger")
        return redirect(url_for("main.dashboard"))

    user = User.query.get_or_404(user_id)
    username = request.form.get("username")
    email = request.form.get("email")
    password = request.form.get("password")
    phone_number = request.form.get("phone_number").strip() if request.form.get("phone_number") else None
    address = request.form.get("address")
    role_id = request.form.get("role_id")
    department_id = request.form.get("department_id")
    image = request.files.get("image")
    remove_image = request.form.get("remove_image")

    if not username or not email:
        flash("Username and email are required!", "danger")
        return redirect(url_for("users.index", page=request.args.get('page', 1, type=int)))

    if username != user.username and User.query.filter_by(username=username).first():
        flash("Username already exists!", "danger")
        return redirect(url_for("users.index", page=request.args.get('page', 1, type=int)))

    if email != user.email and User.query.filter_by(email=email).first():
        flash("Email already exists!", "danger")
        return redirect(url_for("users.index", page=request.args.get('page', 1, type=int)))

    if phone_number and phone_number != user.phone_number and User.query.filter_by(phone_number=phone_number).first():
        flash("Phone number already exists!", "danger")
        return redirect(url_for("users.index", page=request.args.get('page', 1, type=int)))

    if user.role and user.role.name == 'Admin' and User.query.join(Role).filter(Role.name == 'Admin').count() == 1:
        if role_id and role_id != str(user.role_id):
            flash("Cannot change the role of the sole Admin!", "danger")
            return redirect(url_for("users.index", page=request.args.get('page', 1, type=int)))

    if role_id:
        role = Role.query.get(role_id)
        if role and role.name == 'Admin' and User.query.join(Role).filter(Role.name == 'Admin').count() > 0 and user.role.name != 'Admin':
            flash("Only one Admin is allowed in the system!", "danger")
            return redirect(url_for("users.index", page=request.args.get('page', 1, type=int)))

    user.username = username
    user.email = email
    user.phone_number = phone_number
    user.address = address
    user.role_id = int(role_id) if role_id else None
    user.department_id = int(department_id) if department_id else None

    if password:
        user.set_password(password)

    if remove_image == '1' and user.image and user.image != "default_profile.png":
        try:
            image_path = os.path.join(current_app.root_path, 'static/uploads', user.image)
            if os.path.exists(image_path):
                os.remove(image_path)
            user.image = "default_profile.png"
        except Exception as e:
            flash(f"Error removing image: {str(e)}", "danger")

    if image and allowed_file(image.filename):
        filename = secure_filename(image.filename)
        upload_folder = os.path.join(current_app.root_path, 'static/uploads')
        os.makedirs(upload_folder, exist_ok=True)
        file_path = os.path.join(upload_folder, filename)
        try:
            image.save(file_path)
            user.image = filename
        except Exception as e:
            flash(f"Error uploading image: {str(e)}", "danger")
            return redirect(url_for("users.index", page=request.args.get('page', 1, type=int)))

    try:
        db.session.commit()
        flash("User updated successfully!", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error updating user: {str(e)}", "danger")

    return redirect(url_for("users.index", page=request.args.get('page', 1, type=int)))

@users_bp.route("/delete/<int:user_id>", methods=["POST"])
@login_required
def delete(user_id):
    if not current_user.has_role('Admin'):
        flash("You do not have permission to delete users.", "danger")
        return redirect(url_for("main.dashboard"))

    user = User.query.get_or_404(user_id)

    if user.role and user.role.name == 'Admin' and User.query.join(Role).filter(Role.name == 'Admin').count() == 1:
        flash("Cannot delete the sole Admin!", "danger")
        return redirect(url_for("users.index", page=request.args.get('page', 1, type=int)))

    if user.id == current_user.id:
        flash("You cannot delete your own account!", "danger")
        return redirect(url_for("users.index", page=request.args.get('page', 1, type=int)))

    try:
        if user.image and user.image != "default_profile.png":
            image_path = os.path.join(current_app.root_path, 'static/uploads', user.image)
            if os.path.exists(image_path):
                os.remove(image_path)
        db.session.delete(user)
        db.session.commit()
        flash("User deleted successfully!", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error deleting user: {str(e)}", "danger")
    return redirect(url_for("users.index", page=request.args.get('page', 1, type=int)))