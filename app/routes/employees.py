from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file
from flask_login import login_required, current_user
from app import db
from app.models.employees import Employee
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from docxtpl import DocxTemplate
import io
import os
import zipfile
from markupsafe import Markup
from sqlalchemy import or_
import re

employees_bp = Blueprint('employees', __name__)

def build_context(employee):
    """Build context dictionary for DOCX template rendering with improved date formatting."""
    def format_date(date_obj):
        if not date_obj:
            return ''
        day = date_obj.day
        suffix = 'th' if 11 <= day % 100 <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
        superscripts = {"st": "ˢᵗ", "nd": "ⁿᵈ", "rd": "ʳᵈ", "th": "ᵗʰ"}
        return f"{day}{superscripts[suffix]} {date_obj.strftime('%B %Y')}"

    def format_amount(amount):
        try:
            amount = float(amount)
            return str(int(amount)) if amount == int(amount) else f"{amount:.2f}"
        except Exception:
            return ''

    try:
        context = employee.to_dict()
        context['start_date'] = format_date(employee.start_date)
        context['end_date'] = format_date(employee.end_date)
        context['employer_signature_date'] = format_date(employee.employer_signature_date)
        context['employee_signature_date'] = format_date(employee.employee_signature_date)
        context['salary_amount'] = format_amount(employee.salary_amount)
        context['medical_allowance'] = format_amount(employee.medical_allowance)
        context['child_education_allowance'] = format_amount(employee.child_education_allowance)
        context['delivery_benefit'] = format_amount(employee.delivery_benefit)
        context['delivery_benefit_miscarriage'] = format_amount(getattr(employee, 'delivery_benefit_miscarriage', None))
        context['death_benefit'] = format_amount(employee.death_benefit)
        context['severance_percentage'] = f"{float(employee.severance_percentage):.2f}%"
        return context
    except Exception as e:
        flash(f'Error building context for DOCX: {str(e)}', 'danger')
        return {}

def sanitize_filename(name):
    """Sanitize employee name for use in file names by replacing invalid characters."""
    # Replace invalid characters with underscores
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', name)
    # Remove leading/trailing whitespace and periods
    sanitized = sanitized.strip().strip('.')
    # Ensure the name is not empty; use a fallback if necessary
    return sanitized if sanitized else 'Employee'

@employees_bp.route('/')
@login_required
def index():
    try:
        search_query = request.args.get('search', '').strip()
        sort_order = request.args.get('sort', 'created_at_desc')
        entries_per_page = int(request.args.get('entries', 10))
        page = int(request.args.get('page', 1))
        if entries_per_page <= 0:
            entries_per_page = 10

        query = Employee.query.filter_by(deleted_at=None)
        if search_query:
            query = query.filter(
                or_(
                    Employee.employee_name.ilike(f'%{search_query}%'),
                    Employee.position_title.ilike(f'%{search_query}%')
                )
            )

        sort_options = {
            'employee_name_asc': Employee.employee_name.asc(),
            'employee_name_desc': Employee.employee_name.desc(),
            'start_date_asc': Employee.start_date.asc(),
            'start_date_desc': Employee.start_date.desc(),
            'created_at_desc': Employee.created_at.desc()
        }
        query = query.order_by(sort_options.get(sort_order, Employee.created_at.desc()))

        pagination = query.paginate(page=page, per_page=entries_per_page, error_out=False)
        employees = pagination.items
        total_employees = query.count()
        is_admin = current_user.has_role('Admin') if hasattr(current_user, 'has_role') else False

        return render_template(
            'employees/index.html',
            employees=employees,
            pagination=pagination,
            search_query=search_query,
            sort_order=sort_order,
            entries_per_page=entries_per_page,
            total_employees=total_employees,
            is_admin=is_admin
        )
    except Exception as e:
        flash(f'Error loading employee list: {str(e)}', 'danger')
        return render_template(
            'employees/index.html',
            employees=[],
            pagination=None,
            search_query='',
            sort_order='created_at_desc',
            entries_per_page=10,
            total_employees=0,
            is_admin=False
        )

@employees_bp.route('/create', methods=['GET', 'POST'])
@login_required
def create():
    form_data = {}
    if request.method == 'POST':
        try:
            contract_no = request.form['contract_no'].strip()
            start_date_str = request.form['start_date']
            duration_months = int(request.form['duration_months'])
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
            end_date = start_date + relativedelta(months=duration_months)

            new_employee = Employee(
                contract_no=contract_no,
                contract_type=request.form['contract_type'].strip(),
                employee_name=request.form['employee_name'].strip(),
                employee_address=request.form['employee_address'].strip(),
                employee_tel=request.form['employee_tel'].strip(),
                employee_email=request.form['employee_email'].strip(),
                position_title=request.form['position_title'].strip(),
                start_date=start_date,
                end_date=end_date,
                working_hours=request.form['working_hours'].strip(),
                salary_amount=float(request.form['salary_amount']) if request.form['salary_amount'] else 0.0,
                salary_grade=request.form['salary_grade'].strip(),
                medical_allowance=float(request.form['medical_allowance']) if request.form['medical_allowance'] else 150.00,
                child_education_allowance=float(request.form['child_education_allowance']) if request.form['child_education_allowance'] else 60.00,
                delivery_benefit=float(request.form['delivery_benefit']) if request.form['delivery_benefit'] else 200.00,
                delivery_benefit_miscarriage=float(request.form.get('delivery_benefit_miscarriage', 200.00)),
                death_benefit=float(request.form['death_benefit']) if request.form['death_benefit'] else 200.00,
                severance_percentage=float(request.form['severance_percentage']) if request.form['severance_percentage'] else 8.33,
                thirteenth_month_salary=request.form.get('thirteenth_month_salary') == 'on',
                organization_name=request.form['organization_name'].strip(),
                representative_name=request.form['representative_name'].strip(),
                representative_title=request.form.get('representative_title', 'Executive Director').strip(),
                organization_address=request.form['organization_address'].strip(),
                organization_tel=request.form['organization_tel'].strip(),
                organization_fax=request.form['organization_fax'].strip(),
                organization_email=request.form['organization_email'].strip(),
                employer_signature_name=request.form['employer_signature_name'].strip(),
                employee_signature_name=request.form['employee_signature_name'].strip(),
                employer_signature_date=start_date,
                employee_signature_date=start_date
            )

            # Generate salary_amount_words from salary_amount
            new_employee.generate_salary_in_words()

            # Validate client-provided salary_amount_words
            client_salary_words = request.form.get('salary_amount_words', '').strip()
            if client_salary_words and client_salary_words != new_employee.salary_amount_words:
                flash('Warning: Client-provided salary amount in words does not match server-generated value. Using server-generated value.', 'warning')

            db.session.add(new_employee)
            db.session.commit()
            flash('Employee record created successfully!', 'success')
            return redirect(url_for('employees.index'))

        except ValueError as e:
            db.session.rollback()
            flash(f'Invalid input: {str(e)}', 'danger')
            form_data = dict(request.form)
            form_data['thirteenth_month_salary'] = request.form.get('thirteenth_month_salary') == 'on'
        except Exception as e:
            db.session.rollback()
            flash(f'Error creating employee record: {str(e)}', 'danger')
            form_data = dict(request.form)
            form_data['thirteenth_month_salary'] = request.form.get('thirteenth_month_salary') == 'on'

    return render_template('employees/create.html', form_data=form_data)

@employees_bp.route('/<string:id>')
@login_required
def view(id):
    try:
        employee = Employee.query.filter_by(id=id, deleted_at=None).first_or_404()
        return render_template('employees/view.html', employee=employee)
    except Exception as e:
        flash(f'Error viewing employee details: {str(e)}', 'danger')
        return redirect(url_for('employees.index'))

@employees_bp.route('/update/<string:id>', methods=['GET', 'POST'])
@login_required
def update(id):
    try:
        employee = Employee.query.filter_by(id=id, deleted_at=None).first_or_404()
        form_data = {}

        if request.method == 'POST':
            try:
                start_date_str = request.form['start_date']
                start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
                duration_months = int(request.form['duration_months'])
                end_date = start_date + relativedelta(months=duration_months)

                # Update all string fields
                string_fields = [
                    'contract_no', 'contract_type', 'employee_name', 'employee_address',
                    'employee_tel', 'employee_email', 'position_title', 'working_hours',
                    'salary_grade', 'organization_name', 'representative_name', 'representative_title',
                    'organization_address', 'organization_tel', 'organization_fax', 'organization_email',
                    'employer_signature_name', 'employee_signature_name'
                ]
                for field in string_fields:
                    setattr(employee, field, request.form.get(field, '').strip())

                # Update numeric fields safely with defaults
                employee.salary_amount = float(request.form.get('salary_amount', 0.0))
                employee.medical_allowance = float(request.form.get('medical_allowance', 150.00))
                employee.child_education_allowance = float(request.form.get('child_education_allowance', 60.00))
                employee.delivery_benefit = float(request.form.get('delivery_benefit', 200.00))
                employee.delivery_benefit_miscarriage = float(request.form.get('delivery_benefit_miscarriage', 200.00))
                employee.death_benefit = float(request.form.get('death_benefit', 200.00))
                employee.severance_percentage = float(request.form.get('severance_percentage', 8.33))

                # Update boolean fields
                employee.thirteenth_month_salary = request.form.get('thirteenth_month_salary') == 'on'

                # Update dates
                employee.start_date = start_date
                employee.end_date = end_date
                employee.employer_signature_date = start_date
                employee.employee_signature_date = start_date

                # Generate salary_amount_words from salary_amount
                employee.generate_salary_in_words()

                # Validate client-provided salary_amount_words
                client_salary_words = request.form.get('salary_amount_words', '').strip()
                if client_salary_words and client_salary_words != employee.salary_amount_words:
                    flash('Warning: Client-provided salary amount in words does not match server-generated value. Using server-generated value.', 'warning')

                db.session.commit()
                flash('Employee record updated successfully!', 'success')
                return redirect(url_for('employees.index'))

            except ValueError as e:
                db.session.rollback()
                flash(f'Invalid input: {str(e)}', 'danger')
                form_data = dict(request.form)
                form_data['thirteenth_month_salary'] = request.form.get('thirteenth_month_salary') == 'on'
            except Exception as e:
                db.session.rollback()
                flash(f'Error updating employee record: {str(e)}', 'danger')
                form_data = dict(request.form)
                form_data['thirteenth_month_salary'] = request.form.get('thirteenth_month_salary') == 'on'

        else:
            form_data = employee.to_dict()
            if employee.start_date and employee.end_date:
                delta = relativedelta(employee.end_date, employee.start_date)
                form_data['duration_months'] = delta.years * 12 + delta.months
            else:
                form_data['duration_months'] = 12

            form_data['thirteenth_month_salary'] = employee.thirteenth_month_salary

        return render_template('employees/update.html', employee=employee, form_data=form_data)

    except Exception as e:
        flash(f'Error accessing employee record: {str(e)}', 'danger')
        return redirect(url_for('employees.index'))

@employees_bp.route('/delete/<string:id>', methods=['POST'])
@login_required
def delete(id):
    try:
        employee = Employee.query.filter_by(id=id, deleted_at=None).first_or_404()
        employee.deleted_at = datetime.utcnow()
        db.session.commit()
        flash(f'Employee record for {employee.employee_name} deleted successfully!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error deleting employee record: {str(e)}', 'danger')
    return redirect(url_for('employees.index'))

@employees_bp.route('/download/<string:id>')
@login_required
def download_docx(id):
    try:
        employee = Employee.query.filter_by(id=id, deleted_at=None).first_or_404()
        template_path = os.path.join('app', 'static', 'templates', 'Employee_template.docx')
        if not os.path.exists(template_path):
            flash('Template not found. Please ensure Employee_template.docx exists in app/static/templates/', 'danger')
            return redirect(url_for('employees.index'))

        doc = DocxTemplate(template_path)
        doc.render(build_context(employee))
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        filename = f"{employee.contract_no}_{sanitize_filename(employee.employee_name)}.docx"

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        flash(f'Error generating DOCX: {str(e)}', 'danger')
        return redirect(url_for('employees.index'))

@employees_bp.route('/download_all')
@login_required
def download_all_docx():
    try:
        employees = Employee.query.filter_by(deleted_at=None).all()
        if not employees:
            flash("No employee records found to generate.", "warning")
            return redirect(url_for("employees.index"))

        template_path = os.path.join('app', 'static', 'templates', 'Employee_template.docx')
        if not os.path.exists(template_path):
            flash("Template not found.", "danger")
            return redirect(url_for("employees.index"))

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for employee in employees:
                doc = DocxTemplate(template_path)
                doc.render(build_context(employee))
                file_stream = io.BytesIO()
                doc.save(file_stream)
                file_stream.seek(0)
                filename = f"{employee.contract_no}_{sanitize_filename(employee.employee_name)}.docx"
                zip_file.writestr(filename, file_stream.read())

        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=f"All_Employee_Contracts_{datetime.now().strftime('%Y%m%d')}.zip",
            mimetype="application/zip"
        )
    except Exception as e:
        flash(f'Error generating ZIP file: {str(e)}', 'danger')
        return redirect(url_for('employees.index'))

@employees_bp.route('/view_docx/<string:id>')
@login_required
def view_docx(id):
    try:
        employee = Employee.query.filter_by(id=id, deleted_at=None).first_or_404()
        template_path = os.path.join('app', 'static', 'templates', 'Employee_template.docx')
        if not os.path.exists(template_path):
            flash("Template not found.", "danger")
            return redirect(url_for('employees.index'))

        doc = DocxTemplate(template_path)
        doc.render(build_context(employee))
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)

        import mammoth
        result = mammoth.convert_to_html(output)
        html_content = result.value

        return render_template("employees/view_docx.html", html_content=Markup(html_content), employee=employee)
    except Exception as e:
        flash(f'Error rendering contract preview: {str(e)}', 'danger')
        return redirect(url_for('employees.index'))