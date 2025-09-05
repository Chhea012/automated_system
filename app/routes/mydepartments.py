import os
from flask import Blueprint, render_template, request, redirect, url_for, flash, current_app, jsonify
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename
from app import db
from app.models.user import User
from app.models.department import Department
from app.models.role import Role

mydepartments_bp = Blueprint("mydepartments", __name__)

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@mydepartments_bp.route("/")
@login_required
def index():
    if not current_user.department_id:
        flash("You are not assigned to any department.", "warning")
        return redirect(url_for("main.dashboard"))

    if not (current_user.has_role('Admin') or current_user.has_role('Manager')):
        flash("You do not have permission to access this page.", "danger")
        return redirect(url_for("main.dashboard"))

    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '', type=str).strip()
    sort = request.args.get('sort', 'username_asc', type=str)
    per_page = request.args.get('per_page', 7, type=int)

    allowed_per_page = [7, 10, 25, 50]
    if per_page not in allowed_per_page:
        per_page = 7

    query = User.query.filter_by(department_id=current_user.department_id)

    if search:
        query = query.filter(
            (User.username.ilike(f'%{search}%')) | 
            (User.email.ilike(f'%{search}%'))
        )

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

    department = Department.query.get(current_user.department_id)
    total_users = query.count()

    roles = Role.query.filter_by(name='Employee').all()  # Only Employee role for Managers
    departments = [department]  # Only the Manager's department

    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return jsonify({
            'users': [
                {
                    'id': user.id,
                    'username': user.username or 'N/A',
                    'email': user.email or 'N/A',
                    'phone_number': user.phone_number or 'N/A',
                    'address': user.address or 'N/A',
                    'role_name': user.role.name if user.role else 'N/A',
                    'department_name': user.department.name if user.department else 'N/A',
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
        "mydepartments/index.html",
        users=pagination.items,
        pagination=pagination,
        department=department,
        total_users=total_users,
        search=search,
        sort=sort,
        per_page=per_page,
        roles=roles,
        departments=departments
    )

@mydepartments_bp.route("/<int:user_id>/json")
@login_required
def get_user(user_id):
    if not (current_user.has_role('Admin') or current_user.has_role('Manager')):
        return jsonify({'error': 'Unauthorized'}), 403
    user = User.query.get_or_404(user_id)
    if user.department_id != current_user.department_id and not current_user.has_role('Admin'):
        return jsonify({'error': 'Unauthorized'}), 403
    return jsonify({
        'id': user.id,
        'username': user.username,
        'email': user.email,
        'phone_number': user.phone_number,
        'address': user.address,
        'role_id': user.role_id,
        'department_id': user.department_id,
        'image_url': user.get_image_url()
    })

@mydepartments_bp.route("/create", methods=["POST"])
@login_required
def create():
    if not current_user.has_role('Manager') and not current_user.has_role('Admin'):
        flash("You do not have permission to create users.", "danger")
        return redirect(url_for("mydepartments.index", page=1))

    username = request.form.get("username")
    email = request.form.get("email")
    password = request.form.get("password")
    phone_number = request.form.get("phone_number")
    address = request.form.get("address")
    role_id = request.form.get("role_id")
    department_id = current_user.department_id  # Always set to Manager's department
    image = request.files.get("image")

    if not username or not email or not password:
        flash("Username, email, and password are required!", "danger")
        return redirect(url_for("mydepartments.index", page=1))

    if User.query.filter_by(username=username).first():
        flash("Username already exists!", "danger")
        return redirect(url_for("mydepartments.index", page=1))

    if User.query.filter_by(email=email).first():
        flash("Email already exists!", "danger")
        return redirect(url_for("mydepartments.index", page=1))

    # Ensure role_id is for Employee only for Managers
    role = Role.query.get(role_id)
    if not current_user.has_role('Admin') and role and role.name != 'Employee':
        flash("You can only assign the Employee role.", "danger")
        return redirect(url_for("mydepartments.index", page=1))

    new_user = User(username=username, email=email)
    new_user.set_password(password)
    new_user.phone_number = phone_number
    new_user.address = address
    new_user.role_id = int(role_id) if role_id else None
    new_user.department_id = department_id

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
            return redirect(url_for("mydepartments.index", page=1))

    try:
        db.session.add(new_user)
        db.session.commit()
        flash("User created successfully!", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error creating user: {str(e)}", "danger")

    return redirect(url_for("mydepartments.index", page=1))

@mydepartments_bp.route("/update/<int:user_id>", methods=["POST"])
@login_required
def update(user_id):
    if not current_user.has_role('Manager') and not current_user.has_role('Admin'):
        flash("You do not have permission to update users.", "danger")
        return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))

    user = User.query.get_or_404(user_id)
    if user.department_id != current_user.department_id and not current_user.has_role('Admin'):
        flash("You do not have permission to update this user.", "danger")
        return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))

    username = request.form.get("username")
    email = request.form.get("email")
    password = request.form.get("password")
    phone_number = request.form.get("phone_number")
    address = request.form.get("address")
    role_id = request.form.get("role_id")
    department_id = current_user.department_id  # Always set to Manager's department
    image = request.files.get("image")
    remove_image = request.form.get("remove_image")

    if not username or not email:
        flash("Username and email are required!", "danger")
        return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))

    if username != user.username and User.query.filter_by(username=username).first():
        flash("Username already exists!", "danger")
        return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))

    if email != user.email and User.query.filter_by(email=email).first():
        flash("Email already exists!", "danger")
        return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))

    # Ensure role_id is for Employee only for Managers
    role = Role.query.get(role_id)
    if not current_user.has_role('Admin') and role and role.name != 'Employee':
        flash("You can only assign the Employee role.", "danger")
        return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))

    user.username = username
    user.email = email
    user.phone_number = phone_number
    user.address = address
    user.role_id = int(role_id) if role_id else None
    user.department_id = department_id

    if password:
        user.set_password(password)

    if remove_image == '1' and user.image:
        try:
            image_path = os.path.join(current_app.root_path, 'static/uploads', user.image)
            if os.path.exists(image_path):
                os.remove(image_path)
            user.image = None
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
            return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))

    try:
        db.session.commit()
        flash("User updated successfully!", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error updating user: {str(e)}", "danger")

    return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))

@mydepartments_bp.route("/delete/<int:user_id>", methods=["POST"])
@login_required
def delete(user_id):
    if not current_user.has_role('Manager') and not current_user.has_role('Admin'):
        flash("You do not have permission to delete users.", "danger")
        return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))

    user = User.query.get_or_404(user_id)
    if user.department_id != current_user.department_id and not current_user.has_role('Admin'):
        flash("You do not have permission to delete this user.", "danger")
        return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))

    if current_user.id == user_id and not current_user.has_role('Admin'):
        flash("You cannot delete your own account.", "danger")
        return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))

    try:
        if user.image:
            image_path = os.path.join(current_app.root_path, 'static/uploads', user.image)
            if os.path.exists(image_path):
                os.remove(image_path)
        db.session.delete(user)
        db.session.commit()
        flash("User deleted successfully!", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error deleting user: {str(e)}", "danger")
    return redirect(url_for("mydepartments.index", page=request.args.get('page', 1, type=int)))