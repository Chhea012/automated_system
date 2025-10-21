from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file
from flask_login import login_required, current_user
from app import db
from app.models.interns import Intern
from datetime import datetime
from dateutil.relativedelta import relativedelta
from docxtpl import DocxTemplate
import io
import os
import zipfile
from markupsafe import Markup
import mammoth

interns_bp = Blueprint('interns', __name__)

# -------------------------------
# 📝 Helper: Build context for DOCX
# -------------------------------
def build_context(intern):
    def format_date(date):
        return date.strftime('%d %B %Y') if date else ''
    def format_allowance(amount):
        return str(int(amount)) if amount == int(amount) else f"{amount:.2f}"

    context = intern.to_dict()
    context['start_date'] = format_date(intern.start_date)
    context['end_date'] = format_date(intern.end_date)
    context['full_time_period'] = f"Full Time from {context['start_date']} to {context['end_date']}"
    context['allowance_amount'] = format_allowance(float(intern.allowance_amount))
    return context

# -------------------------------
# 📄 List Interns
# -------------------------------
@interns_bp.route('/')
@login_required
def index():
    search_query = request.args.get('search', '')
    sort_order = request.args.get('sort', 'created_at_desc')
    entries_per_page = int(request.args.get('entries', 10))
    page = int(request.args.get('page', 1))

    query = Intern.query.filter_by(deleted_at=None)
    if search_query:
        query = query.filter(
            (Intern.intern_name.ilike(f'%{search_query}%')) |
            (Intern.intern_role.ilike(f'%{search_query}%'))
        )

    if sort_order == 'intern_name_asc':
        query = query.order_by(Intern.intern_name.asc())
    elif sort_order == 'intern_name_desc':
        query = query.order_by(Intern.intern_name.desc())
    elif sort_order == 'start_date_asc':
        query = query.order_by(Intern.start_date.asc())
    elif sort_order == 'start_date_desc':
        query = query.order_by(Intern.start_date.desc())
    else:
        query = query.order_by(Intern.created_at.desc())

    pagination = query.paginate(page=page, per_page=entries_per_page, error_out=False)
    interns = pagination.items
    total_interns = query.count()

    return render_template(
        'interns/index.html',
        interns=interns,
        pagination=pagination,
        search_query=search_query,
        sort_order=sort_order,
        entries_per_page=entries_per_page,
        total_interns=total_interns,
        is_admin=current_user.has_role('Admin') if hasattr(current_user, 'has_role') else False
    )

# -------------------------------
# ➕ Create Intern
# -------------------------------
@interns_bp.route('/create', methods=['GET', 'POST'])
@login_required
def create():
    form_data = {'supervisor_info': {'title': '', 'name': ''}}

    if request.method == 'POST':
        try:
            start_date = datetime.strptime(request.form['start_date'], '%Y-%m-%d')
            duration_months = int(request.form['duration'].split()[0])
            end_date = start_date + relativedelta(months=duration_months)
            allowance = float(request.form['allowance_amount']) if request.form['allowance_amount'] else 0.0
            has_nssf = request.form.get('has_nssf') == 'on'

            new_intern = Intern(
                intern_name=request.form['intern_name'],  # Full name including title
                intern_role=request.form['intern_role'],
                intern_address=request.form['intern_address'],
                intern_phone=request.form['intern_phone'],
                intern_email=request.form['intern_email'],
                start_date=start_date,
                duration=request.form['duration'],
                end_date=end_date,
                working_hours=request.form['working_hours'],
                allowance_amount=allowance,
                has_nssf=has_nssf,
                supervisor_info={
                    'title': request.form['supervisor_title'],
                    'name': request.form['supervisor_name']
                },
                employer_representative_name=request.form['employer_representative_name'],
                employer_representative_title=request.form['employer_representative_title'],
                employer_address=request.form['employer_address'],
                employer_phone=request.form['employer_phone'],
                employer_fax=request.form['employer_fax'],
                employer_email=request.form['employer_email']
            )
            db.session.add(new_intern)
            db.session.commit()
            flash('Intern record created successfully!', 'success')
            return redirect(url_for('interns.index'))
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
    intern = Intern.query.filter_by(id=id, deleted_at=None).first_or_404()
    return render_template('interns/view.html', intern=intern)

# -------------------------------
# ✏️ Update Intern
# -------------------------------
@interns_bp.route('/update/<string:id>', methods=['GET', 'POST'])
@login_required
def update(id):
    intern = Intern.query.filter_by(id=id, deleted_at=None).first_or_404()
    form_data = intern.to_dict()

    if request.method == 'POST':
        try:
            intern.intern_name = request.form['intern_name']
            intern.intern_role = request.form['intern_role']
            intern.intern_address = request.form['intern_address']
            intern.intern_phone = request.form['intern_phone']
            intern.intern_email = request.form['intern_email']
            intern.start_date = datetime.strptime(request.form['start_date'], '%Y-%m-%d')
            duration_months = int(request.form['duration'].split()[0])
            intern.duration = request.form['duration']
            intern.end_date = intern.start_date + relativedelta(months=duration_months)
            intern.working_hours = request.form['working_hours']
            intern.allowance_amount = float(request.form['allowance_amount']) if request.form['allowance_amount'] else 0.0
            intern.has_nssf = request.form.get('has_nssf') == 'on'
            intern.supervisor_info = {
                'title': request.form['supervisor_title'],
                'name': request.form['supervisor_name']
            }
            intern.employer_representative_name = request.form['employer_representative_name']
            intern.employer_representative_title = request.form['employer_representative_title']
            intern.employer_address = request.form['employer_address']
            intern.employer_phone = request.form['employer_phone']
            intern.employer_fax = request.form['employer_fax']
            intern.employer_email = request.form['employer_email']

            db.session.commit()
            flash('Intern record updated successfully!', 'success')
            return redirect(url_for('interns.index'))
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

# -------------------------------
# 🗑 Delete Intern (Soft)
# -------------------------------
@interns_bp.route('/delete/<string:id>', methods=['POST'])
@login_required
def delete(id):
    intern = Intern.query.filter_by(id=id, deleted_at=None).first_or_404()
    try:
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
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

# -------------------------------
# 🆕 Download All DOCX in ZIP
# -------------------------------
@interns_bp.route('/download_all')
@login_required
def download_all_docx():
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
    return send_file(zip_buffer, as_attachment=True,
                     download_name="All_Internship_Agreements.zip",
                     mimetype="application/zip")

# -------------------------------
# 🆕 View DOCX as HTML
# -------------------------------
@interns_bp.route('/view_docx/<string:id>')
@login_required
def view_docx(id):
    """Render the intern DOCX template as HTML for preview."""
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

    # Convert DOCX to HTML for preview
    result = mammoth.convert_to_html(output)
    html_content = result.value

    return render_template("interns/view_docx.html", html_content=Markup(html_content), intern=intern)
