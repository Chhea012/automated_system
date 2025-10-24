from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file
from flask_login import login_required, current_user
from app import db
from app.models.interns import Intern
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from docxtpl import DocxTemplate
import io
import os
import zipfile
from markupsafe import Markup
import mammoth
from sqlalchemy import func
import pandas as pd
import ast

interns_bp = Blueprint('interns', __name__)
# -------------------------------
# 📝 Helper: Build context for DOCX
# -------------------------------
def build_context(intern):
    """Build context dictionary for DOCX template rendering."""
    def format_date(date):
        return date.strftime('%d %B %Y') if date else ''
    def format_allowance(amount):
        return str(int(amount)) if amount == int(amount) else f"{amount:.2f}"
    try:
        context = intern.to_dict()
        context['start_date'] = format_date(intern.start_date)
        context['end_date'] = format_date(intern.end_date)
        context['full_time_period'] = f"Full Time from {context['start_date']} to {context['end_date']}"
        context['allowance_amount'] = format_allowance(float(intern.allowance_amount)) if intern.allowance_amount is not None else '0.00'
        return context
    except Exception as e:
        flash(f'Error building context for DOCX: {str(e)}', 'danger')
        return {}
# -------------------------------
# 📄 List Interns
# -------------------------------
@interns_bp.route('/')
@login_required
def index():
    """Display the list of interns with filters, sorting, and pagination."""
    try:
        search_query = request.args.get('search', '').strip()
        sort_order = request.args.get('sort', 'created_at_desc')
        entries_per_page = request.args.get('entries', 10, type=int)
        page = request.args.get('page', 1, type=int)
        # Validate entries_per_page to prevent negative or zero values
        if entries_per_page <= 0:
            entries_per_page = 10
        # Base query for active interns
        query = Intern.query.filter_by(deleted_at=None)
        # Apply search filter
        if search_query:
            query = query.filter(
                (Intern.intern_name.ilike(f'%{search_query}%')) |
                (Intern.intern_role.ilike(f'%{search_query}%'))
            )
        # Apply sorting
        sort_options = {
            'intern_name_asc': Intern.intern_name.asc(),
            'intern_name_desc': Intern.intern_name.desc(),
            'start_date_asc': Intern.start_date.asc(),
            'start_date_desc': Intern.start_date.desc(),
            'created_at_desc': Intern.created_at.desc()
        }
        query = query.order_by(sort_options.get(sort_order, Intern.created_at.desc()))
        # Pagination
        pagination = query.paginate(page=page, per_page=entries_per_page, error_out=False)
        interns = pagination.items
        total_interns = query.count()
        # Count interns by supervisor title
        supervisor_counts = db.session.query(
            Intern.supervisor_info['title'].label('title'),
            func.count().label('count')
        ).filter(
            Intern.deleted_at == None
        ).group_by(
            Intern.supervisor_info['title']
        ).all()
        supervisor_counts = {row.title: row.count for row in supervisor_counts if row.title}
        # Check admin role
        is_admin = current_user.has_role('Admin') if hasattr(current_user, 'has_role') else False
        return render_template(
            'interns/index.html',
            interns=interns,
            pagination=pagination,
            search_query=search_query,
            sort_order=sort_order,
            entries_per_page=entries_per_page,
            total_interns=total_interns,
            is_admin=is_admin,
            supervisor_counts=supervisor_counts
        )
    except Exception as e:
        flash(f'Error loading intern list: {str(e)}', 'danger')
        # Create a dummy pagination object to avoid template errors
        class DummyPagination:
            def __init__(self):
                self.items = []
                self.page = 1
                self.pages = 1
                self.total = 0
                self.has_prev = False
                self.has_next = False
                self.prev_num = 1
                self.next_num = 1
            def iter_pages(self):
                return []
        return render_template(
            'interns/index.html',
            interns=[],
            pagination=DummyPagination(),
            search_query='',
            sort_order='created_at_desc',
            entries_per_page=10,
            total_interns=0,
            is_admin=False,
            supervisor_counts={}
        )
# -------------------------------
# 📥 Import Interns from Excel
# -------------------------------
@interns_bp.route('/import', methods=['POST'])
@login_required
def import_excel():
    """Import interns from an uploaded Excel file."""
    try:
        if 'file' not in request.files:
            flash('No file part', 'danger')
            return redirect(url_for('interns.index'))
        file = request.files['file']
        if file.filename == '':
            flash('No selected file', 'danger')
            return redirect(url_for('interns.index'))
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            flash('Only Excel files (.xlsx, .xls) are allowed', 'danger')
            return redirect(url_for('interns.index'))
        # Read the Excel file (assuming first sheet, headers in first row)
        df = pd.read_excel(file)
        imported_count = 0
        for _, row in df.iterrows():
            try:
                # Parse start_date (handle serial, datetime, or string)
                start_date_raw = row['Start Date']
                if isinstance(start_date_raw, (int, float)):
                    start_date = datetime(1899, 12, 30) + timedelta(days=start_date_raw)
                elif isinstance(start_date_raw, datetime):
                    start_date = start_date_raw
                elif isinstance(start_date_raw, str):
                    start_date = datetime.strptime(start_date_raw, '%Y-%m-%d')
                else:
                    continue  # Skip invalid row
                # Parse duration and calculate end_date
                duration = str(row['Duration']).strip()
                duration_months = int(duration.split()[0]) if duration else 0
                end_date = start_date + relativedelta(months=duration_months)
                # Parse allowance
                allowance_amount = float(row['Allowance (USD)']) if pd.notna(row['Allowance (USD)']) else 0.0
                # Parse has_nssf
                has_nssf_raw = row['Has NSSF']
                has_nssf = False
                if isinstance(has_nssf_raw, bool):
                    has_nssf = has_nssf_raw
                elif isinstance(has_nssf_raw, str):
                    has_nssf = has_nssf_raw.lower() == 'true'
                # Parse supervisor_info (assume string representation of dict)
                supervisor_raw = row['Supervisor Info']
                supervisor_info = {}
                if isinstance(supervisor_raw, str):
                    supervisor_info = ast.literal_eval(supervisor_raw)
                elif isinstance(supervisor_raw, dict):
                    supervisor_info = supervisor_raw
                # Create new intern
                new_intern = Intern(
                    intern_name=str(row['Intern Name']).strip() if pd.notna(row['Intern Name']) else '',
                    intern_role=str(row['Role']).strip() if pd.notna(row['Role']) else '',
                    intern_address=str(row['Address']).strip() if pd.notna(row['Address']) else '',
                    intern_phone=str(row['Phone']).strip() if pd.notna(row['Phone']) else '',
                    intern_email=str(row['Email']).strip() if pd.notna(row['Email']) else '',
                    start_date=start_date,
                    duration=duration,
                    end_date=end_date,
                    working_hours=str(row['Working Hours']).strip() if pd.notna(row['Working Hours']) else '8:00am – 5:00pm, Monday to Friday',
                    allowance_amount=allowance_amount,
                    has_nssf=has_nssf,
                    supervisor_info=supervisor_info,
                    employer_representative_name=str(row['Employer Representative']).strip() if pd.notna(row['Employer Representative']) else 'Mr. Soeung Saroeun',
                    employer_representative_title=str(row['Title']).strip() if pd.notna(row['Title']) else 'Executive Director',
                    employer_address=str(row['Employer Address']).strip() if pd.notna(row['Employer Address']) else '#9-11, St. 476, Sangkat ToulTompong I, Khan Chamka Morn, Phnom Penh, Cambodia',
                    employer_phone=str(row['Employer Phone']).strip() if pd.notna(row['Employer Phone']) else '023 214 429',
                    employer_fax=str(row['Employer Fax']).strip() if pd.notna(row['Employer Fax']) else '023 994 063',
                    employer_email=str(row['Employer Email']).strip() if pd.notna(row['Employer Email']) else 'info@ngoforum.org.kh'
                )
                db.session.add(new_intern)
                imported_count += 1
            except Exception as row_error:
                db.session.rollback()
                flash(f'Error importing row {_:}: {str(row_error)}', 'warning')
                continue
        db.session.commit()
        flash(f'Successfully imported {imported_count} interns!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error importing Excel file: {str(e)}', 'danger')
    return redirect(url_for('interns.index'))
# -------------------------------
# ➕ Create Intern
# -------------------------------
@interns_bp.route('/create', methods=['GET', 'POST'])
@login_required
def create():
    """Create a new intern record."""
    form_data = {'supervisor_info': {'title': '', 'name': ''}}
    if request.method == 'POST':
        try:
            start_date = datetime.strptime(request.form['start_date'], '%Y-%m-%d')
            duration_months = int(request.form['duration'].split()[0])
            end_date = start_date + relativedelta(months=duration_months)
            allowance = float(request.form['allowance_amount']) if request.form['allowance_amount'] else 0.0
            has_nssf = request.form.get('has_nssf') == 'on'
            new_intern = Intern(
                intern_name=request.form['intern_name'].strip(),
                intern_role=request.form['intern_role'].strip(),
                intern_address=request.form['intern_address'].strip(),
                intern_phone=request.form['intern_phone'].strip(),
                intern_email=request.form['intern_email'].strip(),
                start_date=start_date,
                duration=request.form['duration'].strip(),
                end_date=end_date,
                working_hours=request.form['working_hours'].strip(),
                allowance_amount=allowance,
                has_nssf=has_nssf,
                supervisor_info={
                    'title': request.form['supervisor_title'].strip(),
                    'name': request.form['supervisor_name'].strip()
                },
                employer_representative_name=request.form['employer_representative_name'].strip(),
                employer_representative_title=request.form['employer_representative_title'].strip(),
                employer_address=request.form['employer_address'].strip(),
                employer_phone=request.form['employer_phone'].strip(),
                employer_fax=request.form['employer_fax'].strip(),
                employer_email=request.form['employer_email'].strip()
            )
            db.session.add(new_intern)
            db.session.commit()
            flash('Intern record created successfully!', 'success')
            return redirect(url_for('interns.index'))
        except ValueError as e:
            db.session.rollback()
            flash(f'Invalid input: {str(e)}', 'danger')
            form_data = request.form.to_dict()
            form_data['supervisor_info'] = {
                'title': request.form.get('supervisor_title', ''),
                'name': request.form.get('supervisor_name', '')
            }
            form_data['has_nssf'] = request.form.get('has_nssf') == 'on'
        except Exception as e:
            db.session.rollback()
            flash(f'Error creating intern record: {str(e)}', 'danger')
            form_data = request.form.to_dict()
            form_data['supervisor_info'] = {
                'title': request.form.get('supervisor_title', ''),
                'name': request.form.get('supervisor_name', '')
            }
            form_data['has_nssf'] = request.form.get('has_nssf') == 'on'
    return render_template('interns/create.html', form_data=form_data)
# -------------------------------
# 👁 View Intern Details
# -------------------------------
@interns_bp.route('/<string:id>')
@login_required
def view(id):
    """View details of a specific intern."""
    try:
        intern = Intern.query.filter_by(id=id, deleted_at=None).first_or_404()
        return render_template('interns/view.html', intern=intern)
    except Exception as e:
        flash(f'Error viewing intern details: {str(e)}', 'danger')
        return redirect(url_for('interns.index'))
# -------------------------------
# ✏️ Update Intern
# -------------------------------
@interns_bp.route('/update/<string:id>', methods=['GET', 'POST'])
@login_required
def update(id):
    """Update an existing intern record."""
    try:
        intern = Intern.query.filter_by(id=id, deleted_at=None).first_or_404()
        form_data = intern.to_dict()
        if request.method == 'POST':
            try:
                intern.intern_name = request.form['intern_name'].strip()
                intern.intern_role = request.form['intern_role'].strip()
                intern.intern_address = request.form['intern_address'].strip()
                intern.intern_phone = request.form['intern_phone'].strip()
                intern.intern_email = request.form['intern_email'].strip()
                intern.start_date = datetime.strptime(request.form['start_date'], '%Y-%m-%d')
                duration_months = int(request.form['duration'].split()[0])
                intern.duration = request.form['duration'].strip()
                intern.end_date = intern.start_date + relativedelta(months=duration_months)
                intern.working_hours = request.form['working_hours'].strip()
                intern.allowance_amount = float(request.form['allowance_amount']) if request.form['allowance_amount'] else 0.0
                intern.has_nssf = request.form.get('has_nssf') == 'on'
                intern.supervisor_info = {
                    'title': request.form['supervisor_title'].strip(),
                    'name': request.form['supervisor_name'].strip()
                }
                intern.employer_representative_name = request.form['employer_representative_name'].strip()
                intern.employer_representative_title = request.form['employer_representative_title'].strip()
                intern.employer_address = request.form['employer_address'].strip()
                intern.employer_phone = request.form['employer_phone'].strip()
                intern.employer_fax = request.form['employer_fax'].strip()
                intern.employer_email = request.form['employer_email'].strip()
                db.session.commit()
                flash('Intern record updated successfully!', 'success')
                return redirect(url_for('interns.index'))
            except ValueError as e:
                db.session.rollback()
                flash(f'Invalid input: {str(e)}', 'danger')
                form_data = request.form.to_dict()
                form_data['supervisor_info'] = {
                    'title': request.form.get('supervisor_title', ''),
                    'name': request.form.get('supervisor_name', '')
                }
                form_data['has_nssf'] = request.form.get('has_nssf') == 'on'
            except Exception as e:
                db.session.rollback()
                flash(f'Error updating intern record: {str(e)}', 'danger')
                form_data = request.form.to_dict()
                form_data['supervisor_info'] = {
                    'title': request.form.get('supervisor_title', ''),
                    'name': request.form.get('supervisor_name', '')
                }
                form_data['has_nssf'] = request.form.get('has_nssf') == 'on'
        return render_template('interns/update.html', intern=intern, form_data=form_data)
    except Exception as e:
        flash(f'Error accessing intern record: {str(e)}', 'danger')
        return redirect(url_for('interns.index'))
# -------------------------------
# 🗑 Delete Intern (Soft)
# -------------------------------
@interns_bp.route('/delete/<string:id>', methods=['POST'])
@login_required
def delete(id):
    """Soft delete an intern record."""
    try:
        intern = Intern.query.filter_by(id=id, deleted_at=None).first_or_404()
        intern.deleted_at = datetime.utcnow()
        db.session.commit()
        flash(f'Intern record for {intern.intern_name} deleted successfully!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error deleting intern record: {str(e)}', 'danger')
    return redirect(url_for('interns.index'))
# -------------------------------
# 🧾 Download Single DOCX
# -------------------------------
@interns_bp.route('/download/<string:id>')
@login_required
def download_docx(id):
    """Download a single intern's contract as DOCX."""
    try:
        intern = Intern.query.filter_by(id=id, deleted_at=None).first_or_404()
        template_path = os.path.join('app', 'static', 'templates', 'internship_template.docx')
        if not os.path.exists(template_path):
            flash('Template not found.', 'danger')
            return redirect(url_for('interns.index'))
        doc = DocxTemplate(template_path)
        doc.render(build_context(intern))
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        filename = f"{intern.intern_name.replace(' ', '_')}_Internship_Agreement.docx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        flash(f'Error generating DOCX: {str(e)}', 'danger')
        return redirect(url_for('interns.index'))
# -------------------------------
# 🆕 Download All DOCX in ZIP
# -------------------------------
@interns_bp.route('/download_all')
@login_required
def download_all_docx():
    """Download all intern contracts as a ZIP file."""
    try:
        interns = Intern.query.filter_by(deleted_at=None).all()
        if not interns:
            flash("No intern records found to generate.", "warning")
            return redirect(url_for("interns.index"))
        template_path = os.path.join('app', 'static', 'templates', 'internship_template.docx')
        if not os.path.exists(template_path):
            flash("Template not found.", "danger")
            return redirect(url_for("interns.index"))
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for intern in interns:
                doc = DocxTemplate(template_path)
                doc.render(build_context(intern))
                file_stream = io.BytesIO()
                doc.save(file_stream)
                file_stream.seek(0)
                filename = f"{intern.intern_name.replace(' ', '_')}_Internship_Agreement.docx"
                zip_file.writestr(filename, file_stream.read())
        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name="All_Internship_Agreements.zip",
            mimetype="application/zip"
        )
    except Exception as e:
        flash(f'Error generating ZIP file: {str(e)}', 'danger')
        return redirect(url_for('interns.index'))
# -------------------------------
# 🆕 View DOCX as HTML
# -------------------------------
@interns_bp.route('/view_docx/<string:id>')
@login_required
def view_docx(id):
    """Render the intern DOCX template as HTML for preview."""
    try:
        intern = Intern.query.filter_by(id=id, deleted_at=None).first_or_404()
        template_path = os.path.join('app', 'static', 'templates', 'internship_template.docx')
        if not os.path.exists(template_path):
            flash("Template not found.", "danger")
            return redirect(url_for('interns.index'))
        doc = DocxTemplate(template_path)
        doc.render(build_context(intern))
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        result = mammoth.convert_to_html(output)
        html_content = result.value
        return render_template("interns/view_docx.html", html_content=Markup(html_content), intern=intern)
    except Exception as e:
        flash(f'Error rendering contract preview: {str(e)}', 'danger')
        return redirect(url_for('interns.index'))