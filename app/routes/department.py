from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required
from app import db
from app.models.department import Department

departments_bp = Blueprint("departments", __name__)

# List departments with pagination, sorted by created_at descending
@departments_bp.route("/")
@login_required
def index():
    page = request.args.get('page', 1, type=int)
    pagination = Department.query.order_by(Department.created_at.desc()).paginate(page=page, per_page=7, error_out=False)
    return render_template("departments/index.html", departments=pagination.items, pagination=pagination)

# Create department
@departments_bp.route("/create", methods=["POST"])
@login_required
def create():
    name = request.form.get("name")
    description = request.form.get("description")

    if not name:
        flash("Department name is required!", "danger")
        return redirect(url_for("departments.index", page=1))

    if Department.query.filter_by(name=name).first():
        flash("Department name already exists!", "danger")
        return redirect(url_for("departments.index", page=1))

    new_department = Department(
        name=name,
        description=description
    )

    db.session.add(new_department)
    db.session.commit()
    flash("Department created successfully!", "success")
    return redirect(url_for("departments.index", page=1))

# Update department
@departments_bp.route("/update/<int:department_id>", methods=["POST"])
@login_required
def update(department_id):
    department = Department.query.get_or_404(department_id)
    name = request.form.get("name")
    description = request.form.get("description")

    if not name:
        flash("Department name is required!", "danger")
        return redirect(url_for("departments.index", page=request.args.get('page', 1)))

    if name != department.name and Department.query.filter_by(name=name).first():
        flash("Department name already exists!", "danger")
        return redirect(url_for("departments.index", page=request.args.get('page', 1)))

    department.name = name
    department.description = description

    db.session.commit()
    flash("Department updated successfully!", "success")
    return redirect(url_for("departments.index", page=request.args.get('page', 1)))

# Delete department
@departments_bp.route("/delete/<int:department_id>", methods=["POST"])
@login_required
def delete(department_id):
    department = Department.query.get_or_404(department_id)
    db.session.delete(department)
    db.session.commit()
    flash("Department deleted!", "success")
    return redirect(url_for("departments.index", page=request.args.get('page', 1)))