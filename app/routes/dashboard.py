from flask import Blueprint, render_template
from flask_login import login_required, current_user
from app import db
from app.models.user import User
from app.models.role import Role
from app.models.department import Department

dashboard_bp = Blueprint("dashboard", __name__)

@dashboard_bp.route("/dashboard")
@login_required
def index():
    departments = Department.query.all()
    department_data = []
    for dept in departments:
        users = User.query.filter_by(department_id=dept.id).all()
        user_count = len(users)
        managers = [
            user.username for user in users 
            if user.role and user.role.name == 'Manager'
        ]
        employees = [
            user.username for user in users 
            if user.role and user.role.name == 'Employee'
        ]
        department_data.append({
            'name': dept.name or 'N/A',
            'user_count': user_count,
            'managers': managers,
            'employees': employees
        })
    pie_labels = [dept['name'] for dept in department_data]
    pie_data = [dept['user_count'] for dept in department_data]
    if not department_data:
        pie_labels = ['No Departments']
        pie_data = [0]
    return render_template(
        "dashboard.html",
        user=current_user,
        department_data=department_data,
        pie_labels=pie_labels,
        pie_data=pie_data
    )