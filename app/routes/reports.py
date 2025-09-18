from flask import Blueprint, render_template, request, send_file
from flask_login import login_required
from app import db
from app.models.contract import Contract
from app.models.department import Department
from app.models.user import User  # Import User model explicitly
from datetime import datetime
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

reports_bp = Blueprint('reports', __name__)

@reports_bp.route('/contracts')
@login_required
def contract_report():
    # Filters
    department_id = request.args.get('department_id', 'all')
    search = request.args.get('search', '').strip().lower()
    month_year = request.args.get('month_year', datetime.now().strftime('%B %Y'))
    sort = request.args.get('sort', 'contract_number_asc')
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 10, type=int)

    # Parse month_year for filtering (e.g., 'September 2025' -> year=2025, month=9)
    try:
        month_name, year_str = month_year.split()
        month = datetime.strptime(month_name, '%B').month
        year = int(year_str)
    except (ValueError, AttributeError):
        month = datetime.now().month
        year = datetime.now().year
        month_year = datetime.now().strftime('%B %Y')

    # Base query with left join to handle missing users
    query = Contract.query.filter(Contract.deleted_at == None)\
                          .outerjoin(Contract.user)\
                          .filter(db.extract('year', Contract.created_at) == year,
                                  db.extract('month', Contract.created_at) == month)

    if department_id != 'all':
        query = query.filter(Contract.user.has(department_id=department_id))

    if search:
        query = query.filter(
            (Contract.project_title.ilike(f'%{search}%')) |
            (Contract.party_b_signature_name.ilike(f'%{search}%')) |
            (User.username.ilike(f'%{search}%') & (Contract.user_id != None))
        )

    # Sorting
    if sort == 'contract_number_desc':
        query = query.order_by(Contract.contract_number.desc())
    elif sort == 'project_title_asc':
        query = query.order_by(Contract.project_title.asc())
    elif sort == 'project_title_desc':
        query = query.order_by(Contract.project_title.desc())
    else:
        query = query.order_by(Contract.contract_number.asc())

    pagination = query.paginate(page=page, per_page=per_page, error_out=False)
    contracts = [contract.to_dict() for contract in pagination.items]

    # Totals
    total_contracts = query.count()
    departments = Department.query.all()
    department_totals = {dept.name: Contract.query.filter(Contract.deleted_at == None)
                                               .outerjoin(Contract.user)
                                               .filter(Contract.user.has(department_id=dept.id))
                                               .filter(db.extract('year', Contract.created_at) == year,
                                                       db.extract('month', Contract.created_at) == month).count()
                         for dept in departments}

    # Unique month_years for dropdown (MySQL-compatible)
    unique_months = db.session.query(db.func.distinct(db.func.date_format(Contract.created_at, '%M %Y')))\
                              .filter(Contract.created_at != None).all()
    unique_months = [m[0] for m in unique_months if m[0]]

    return render_template('reports/index.html',
                           contracts=contracts,
                           pagination=pagination,
                           departments=departments,
                           department_id=department_id,
                           search=search,
                           month_year=month_year,
                           sort=sort,
                           per_page=per_page,
                           total_contracts=total_contracts,
                           department_totals=department_totals,
                           unique_months=unique_months)

@reports_bp.route('/export_contracts_excel')
@login_required
def export_contracts_excel():
    department_id = request.args.get('department_id', 'all')
    month_year = request.args.get('month_year', datetime.now().strftime('%B %Y'))
    search = request.args.get('search', '').strip().lower()

    try:
        month_name, year_str = month_year.split()
        month = datetime.strptime(month_name, '%B').month
        year = int(year_str)
    except (ValueError, AttributeError):
        month = datetime.now().month
        year = datetime.now().year

    query = Contract.query.filter(Contract.deleted_at == None)\
                          .outerjoin(Contract.user)\
                          .filter(db.extract('year', Contract.created_at) == year,
                                  db.extract('month', Contract.created_at) == month)

    if department_id != 'all' and department_id != 'current':
        query = query.filter(Contract.user.has(department_id=department_id))

    if search:
        query = query.filter(
            (Contract.project_title.ilike(f'%{search}%')) |
            (Contract.party_b_signature_name.ilike(f'%{search}%')) |
            (User.username.ilike(f'%{search}%') & (Contract.user_id != None))
        )

    contracts = query.all()
    data = [{
        'Number of Contract': c.contract_number,
        'Project Title': c.project_title,
        'Department': c.user.department.name if c.user and c.user.department else 'N/A',
        'Manager': c.user.username if c.user else 'N/A',
        'Contract Date': c.formatted_created_at,
        'Party B': c.party_b_signature_name
    } for c in contracts]

    df = pd.DataFrame(data)

    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Contract Report"

    # Add headers
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # Styling
    header_font = Font(bold=True)
    alignment = Alignment(horizontal='center', vertical='center')
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    fill = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid')

    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = alignment
        cell.border = border
        cell.fill = fill

    # Auto-adjust columns
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    # Add totals row
    total_row = ws.max_row + 1
    ws.cell(row=total_row, column=1, value='Total Contracts').font = Font(bold=True)
    ws.cell(row=total_row, column=2, value=len(contracts)).font = Font(bold=True)

    wb.save(output)
    output.seek(0)

    filename = f"Contract_Report_{month_year.replace(' ', '_')}.xlsx"
    return send_file(output, download_name=filename, as_attachment=True)