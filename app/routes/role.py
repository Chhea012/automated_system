from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
from app import db
from app.models.role import Role
from app.models.permission import Permission

roles_bp = Blueprint("roles", __name__)

@roles_bp.route("/")
@login_required
def index():
    if not current_user.has_role('Admin'):
        flash("You do not have permission to access this page.", "danger")
        return redirect(url_for("main.dashboard"))

    page = request.args.get('page', 1, type=int)
    pagination = Role.query.order_by(Role.created_at.desc()).paginate(page=page, per_page=7, error_out=False)
    permissions = Permission.query.all()
    return render_template("roles/index.html", roles=pagination.items, pagination=pagination, permissions=permissions)

@roles_bp.route("/create", methods=["POST"])
@login_required
def create():
    if not current_user.has_role('Admin'):
        flash("You do not have permission to create roles.", "danger")
        return redirect(url_for("main.dashboard"))

    name = request.form.get("name")
    description = request.form.get("description")
    permission_ids = request.form.getlist("permissions")

    if not name:
        flash("Role name is required!", "danger")
        return redirect(url_for("roles.index", page=1))

    if Role.query.filter_by(name=name).first():
        flash("Role name already exists!", "danger")
        return redirect(url_for("roles.index", page=1))

    new_role = Role(
        name=name,
        description=description
    )

    if permission_ids:
        permissions = Permission.query.filter(Permission.id.in_(permission_ids)).all()
        new_role.permissions = permissions

    db.session.add(new_role)
    db.session.commit()
    flash("Role created successfully!", "success")
    return redirect(url_for("roles.index", page=1))

@roles_bp.route("/update/<int:role_id>", methods=["POST"])
@login_required
def update(role_id):
    if not current_user.has_role('Admin'):
        flash("You do not have permission to update roles.", "danger")
        return redirect(url_for("main.dashboard"))

    role = Role.query.get_or_404(role_id)
    name = request.form.get("name")
    description = request.form.get("description")
    permission_ids = request.form.getlist("permissions")

    if not name:
        flash("Role name is required!", "danger")
        return redirect(url_for("roles.index", page=request.args.get('page', 1)))

    if name != role.name and Role.query.filter_by(name=name).first():
        flash("Role name already exists!", "danger")
        return redirect(url_for("roles.index", page=request.args.get('page', 1)))

    role.name = name
    role.description = description

    role.permissions = []
    if permission_ids:
        permissions = Permission.query.filter(Permission.id.in_(permission_ids)).all()
        role.permissions = permissions

    db.session.commit()
    flash("Role updated successfully!", "success")
    return redirect(url_for("roles.index", page=request.args.get('page', 1)))

@roles_bp.route("/delete/<int:role_id>", methods=["POST"])
@login_required
def delete(role_id):
    if not current_user.has_role('Admin'):
        flash("You do not have permission to delete roles.", "danger")
        return redirect(url_for("main.dashboard"))

    role = Role.query.get_or_404(role_id)
    db.session.delete(role)
    db.session.commit()
    flash("Role deleted!", "success")
    return redirect(url_for("roles.index", page=request.args.get('page', 1)))