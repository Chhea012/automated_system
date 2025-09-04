from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
from app import db
from app.models.permission import Permission

permissions_bp = Blueprint("permissions", __name__)

@permissions_bp.route("/")
@login_required
def index():
    if not current_user.has_role('Admin'):
        flash("You do not have permission to access this page.", "danger")
        return redirect(url_for("main.dashboard"))

    page = request.args.get('page', 1, type=int)
    pagination = Permission.query.order_by(Permission.created_at.desc()).paginate(page=page, per_page=7, error_out=False)
    return render_template("permissions/index.html", permissions=pagination.items, pagination=pagination)

@permissions_bp.route("/create", methods=["POST"])
@login_required
def create():
    if not current_user.has_role('Admin'):
        flash("You do not have permission to create permissions.", "danger")
        return redirect(url_for("main.dashboard"))

    name = request.form.get("name")
    description = request.form.get("description")

    if not name:
        flash("Permission name is required!", "danger")
        return redirect(url_for("permissions.index", page=1))

    if Permission.query.filter_by(name=name).first():
        flash("Permission name already exists!", "danger")
        return redirect(url_for("permissions.index", page=1))

    new_permission = Permission(
        name=name,
        description=description
    )

    db.session.add(new_permission)
    db.session.commit()
    flash("Permission created successfully!", "success")
    return redirect(url_for("permissions.index", page=1))

@permissions_bp.route("/update/<int:permission_id>", methods=["POST"])
@login_required
def update(permission_id):
    if not current_user.has_role('Admin'):
        flash("You do not have permission to update permissions.", "danger")
        return redirect(url_for("main.dashboard"))

    permission = Permission.query.get_or_404(permission_id)
    name = request.form.get("name")
    description = request.form.get("description")

    if not name:
        flash("Permission name is required!", "danger")
        return redirect(url_for("permissions.index", page=request.args.get('page', 1)))

    if name != permission.name and Permission.query.filter_by(name=name).first():
        flash("Permission name already exists!", "danger")
        return redirect(url_for("permissions.index", page=request.args.get('page', 1)))

    permission.name = name
    permission.description = description

    db.session.commit()
    flash("Permission updated successfully!", "success")
    return redirect(url_for("permissions.index", page=request.args.get('page', 1)))

@permissions_bp.route("/delete/<int:permission_id>", methods=["POST"])
@login_required
def delete(permission_id):
    if not current_user.has_role('Admin'):
        flash("You do not have permission to delete permissions.", "danger")
        return redirect(url_for("main.dashboard"))

    permission = Permission.query.get_or_404(permission_id)
    db.session.delete(permission)
    db.session.commit()
    flash("Permission deleted!", "success")
    return redirect(url_for("permissions.index", page=request.args.get('page', 1)))