from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file
from flask_login import login_required
from app import db
from app.models.contract import Contract
import uuid
import json
from datetime import datetime
import pandas as pd
from io import BytesIO

contracts_bp = Blueprint('contracts', __name__)

# Helper function to format date
def format_date(iso_date):
    try:
        date = datetime.strptime(iso_date, '%Y-%m-%d')
        day = date.day
        month = date.strftime('%B')
        year = date.year
        suffix = 'th' if 11 <= day % 100 <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
        return f"{day}{suffix} {month} {year}"
    except (ValueError, TypeError):
        return iso_date or 'N/A'

# List contracts with pagination, search, and sorting
@contracts_bp.route('/')
@login_required
def index():
    page = request.args.get('page', 1, type=int)
    search_query = request.args.get('search', '', type=str)
    sort_order = request.args.get('sort', 'asc', type=str)
    entries_per_page = request.args.get('entries', 10, type=int)

    query = Contract.query
    if search_query:
        query = query.filter(Contract.project_title.ilike(f'%{search_query}%'))

    if sort_order == 'asc':
        query = query.order_by(Contract.project_title.asc())
    else:
        query = query.order_by(Contract.project_title.desc())

    pagination = query.paginate(page=page, per_page=entries_per_page, error_out=False)
    contracts = [contract.to_dict() for contract in pagination.items]
    for contract in contracts:
        contract['agreement_start_date_display'] = format_date(contract['agreement_start_date'])
        contract['agreement_end_date_display'] = format_date(contract['agreement_end_date'])
        contract['articles'] = [{'articleNumber': k, 'customSentence': v} for k, v in contract['custom_article_sentences'].items()] if contract['custom_article_sentences'] else []
        contract['total_fee_usd'] = f"{contract['total_fee_usd']:.2f}" if contract['total_fee_usd'] is not None else '0.00'

    return render_template('contracts/index.html', contracts=contracts, pagination=pagination,
                           search_query=search_query, sort_order=sort_order, entries_per_page=entries_per_page)

# Create contract
@contracts_bp.route('/create', methods=['GET', 'POST'])
@login_required
def create():
    if request.method == 'POST':
        try:
            project_title = request.form.get('projectTitle')
            contract_number = request.form.get('contractNumber')
            party_b_signature_name = request.form.get('partyBSignatureName')
            party_b_signature_name_confirm = request.form.get('partyBSignatureNameConfirm')

            if not project_title or not contract_number or not party_b_signature_name:
                flash('Project title, contract number, and Party B signature name are required!', 'danger')
                return redirect(url_for('contracts.create'))

            if party_b_signature_name != party_b_signature_name_confirm:
                flash('Party B Signature Name and Confirmation do not match.', 'danger')
                return redirect(url_for('contracts.create'))

            if Contract.query.filter_by(contract_number=contract_number).first():
                flash('Contract number already exists!', 'danger')
                return redirect(url_for('contracts.create'))

            start_date = request.form.get('agreementStartDate')
            end_date = request.form.get('agreementEndDate')
            if start_date and end_date and datetime.strptime(end_date, '%Y-%m-%d') < datetime.strptime(start_date, '%Y-%m-%d'):
                flash('End date cannot be before start date.', 'danger')
                return redirect(url_for('contracts.create'))

            payment_installments = request.form.getlist('paymentInstallmentDesc')
            payment_installments = [desc for desc in payment_installments if desc.strip()]
            total_percentage = sum(float(desc.split('(')[1].split('%')[0]) for desc in payment_installments if desc and '(' in desc and desc.endswith('%)'))
            if total_percentage > 100:
                flash('Total payment installment percentages cannot exceed 100%.', 'danger')
                return redirect(url_for('contracts.create'))

            articles = []
            article_numbers = set()
            for i, article_number in enumerate(request.form.getlist('articleNumber')):
                custom_sentence = request.form.getlist('customSentence')[i]
                if custom_sentence and article_number not in article_numbers:
                    articles.append({'articleNumber': article_number, 'customSentence': custom_sentence})
                    article_numbers.add(article_number)
                elif custom_sentence and article_number in article_numbers:
                    flash(f'Duplicate article number {article_number} detected.', 'danger')
                    return redirect(url_for('contracts.create'))

            deliverables = request.form.get('deliverables', '').split('\n')
            deliverables = [d.strip() for d in deliverables if d.strip()]

            total_fee_usd = float(request.form.get('totalFeeUSD')) if request.form.get('totalFeeUSD') else None
            tax_percentage = float(request.form.get('taxPercentage')) if request.form.get('taxPercentage') else None

            new_contract = Contract(
                id=str(uuid.uuid4()),
                project_title=project_title,
                contract_number=contract_number,
                organization_name=request.form.get('organizationName', 'The NGO Forum on Cambodia'),
                party_a_name=request.form.get('partyAName', 'Mr. Soeung Saroeun'),
                party_a_position=request.form.get('partyAPosition', 'Executive Director'),
                party_a_address=request.form.get('partyAAddress', '#9-11, Street 476, Sangkat Tuol Tumpoung I, Phnom Penh, Cambodia'),
                party_b_full_name_with_title=party_b_signature_name,
                party_b_address=request.form.get('partyBAddress'),
                party_b_phone=request.form.get('partyBPhone'),
                party_b_email=request.form.get('partyBEmail'),
                registration_number=request.form.get('registrationNumber', '#304 សជណ'),
                registration_date=request.form.get('registrationDate', '07 March 2012'),
                agreement_start_date=start_date,
                agreement_end_date=end_date,
                total_fee_usd=total_fee_usd,
                gross_amount_usd=total_fee_usd,
                tax_percentage=tax_percentage,
                payment_installment_desc='; '.join(payment_installments) if payment_installments else None,
                payment_gross=f"USD {total_fee_usd:.2f}" if total_fee_usd else None,
                payment_net=f"USD {total_fee_usd * (1 - tax_percentage / 100):.2f}" if total_fee_usd and tax_percentage is not None else None,
                workshop_description=request.form.get('workshopDescription'),
                focal_person_a_name=request.form.get('focalPersonAName'),
                focal_person_a_position=request.form.get('focalPersonAPosition'),
                focal_person_a_phone=request.form.get('focalPersonAPhone'),
                focal_person_a_email=request.form.get('focalPersonAEmail'),
                party_a_signature_name=request.form.get('partyASignatureName', 'Mr. SOEUNG Saroeun'),
                party_b_signature_name=party_b_signature_name,
                party_b_position=request.form.get('partyBPosition'),
                total_fee_words=request.form.get('totalFeeWords'),
                title=request.form.get('title'),
                deliverables='; '.join(deliverables) if deliverables else None,
                output_description=request.form.get('outputDescription'),
                custom_article_sentences=json.dumps({article['articleNumber']: article['customSentence'] for article in articles}) if articles else '{}'
            )

            db.session.add(new_contract)
            db.session.commit()
            flash('Contract created successfully!', 'success')
            return redirect(url_for('contracts.index'))
        except Exception as e:
            db.session.rollback()
            flash(f'Error creating contract: {str(e)}', 'danger')
            return redirect(url_for('contracts.create'))

    return render_template('contracts/create.html')

# Update contract
@contracts_bp.route('/update/<string:contract_id>', methods=['GET', 'POST'])
@login_required
def update(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    if request.method == 'POST':
        try:
            project_title = request.form.get('projectTitle')
            contract_number = request.form.get('contractNumber')
            party_b_signature_name = request.form.get('partyBSignatureName')
            party_b_signature_name_confirm = request.form.get('partyBSignatureNameConfirm')

            if not project_title or not contract_number or not party_b_signature_name:
                flash('Project title, contract number, and Party B signature name are required!', 'danger')
                return redirect(url_for('contracts.update', contract_id=contract_id))

            if party_b_signature_name != party_b_signature_name_confirm:
                flash('Party B Signature Name and Confirmation do not match.', 'danger')
                return redirect(url_for('contracts.update', contract_id=contract_id))

            if contract_number != contract.contract_number and Contract.query.filter_by(contract_number=contract_number).first():
                flash('Contract number already exists!', 'danger')
                return redirect(url_for('contracts.update', contract_id=contract_id))

            start_date = request.form.get('agreementStartDate')
            end_date = request.form.get('agreementEndDate')
            if start_date and end_date and datetime.strptime(end_date, '%Y-%m-%d') < datetime.strptime(start_date, '%Y-%m-%d'):
                flash('End date cannot be before start date.', 'danger')
                return redirect(url_for('contracts.update', contract_id=contract_id))

            payment_installments = request.form.getlist('paymentInstallmentDesc')
            payment_installments = [desc for desc in payment_installments if desc.strip()]
            total_percentage = sum(float(desc.split('(')[1].split('%')[0]) for desc in payment_installments if desc and '(' in desc and desc.endswith('%)'))
            if total_percentage > 100:
                flash('Total payment installment percentages cannot exceed 100%.', 'danger')
                return redirect(url_for('contracts.update', contract_id=contract_id))

            articles = []
            article_numbers = set()
            for i, article_number in enumerate(request.form.getlist('articleNumber')):
                custom_sentence = request.form.getlist('customSentence')[i]
                if custom_sentence and article_number not in article_numbers:
                    articles.append({'articleNumber': article_number, 'customSentence': custom_sentence})
                    article_numbers.add(article_number)
                elif custom_sentence and article_number in article_numbers:
                    flash(f'Duplicate article number {article_number} detected.', 'danger')
                    return redirect(url_for('contracts.update', contract_id=contract_id))

            deliverables = request.form.get('deliverables', '').split('\n')
            deliverables = [d.strip() for d in deliverables if d.strip()]

            total_fee_usd = float(request.form.get('totalFeeUSD')) if request.form.get('totalFeeUSD') else None
            tax_percentage = float(request.form.get('taxPercentage')) if request.form.get('taxPercentage') else None

            contract.project_title = project_title
            contract.contract_number = contract_number
            contract.organization_name = request.form.get('organizationName', 'The NGO Forum on Cambodia')
            contract.party_a_name = request.form.get('partyAName', 'Mr. Soeung Saroeun')
            contract.party_a_position = request.form.get('partyAPosition', 'Executive Director')
            contract.party_a_address = request.form.get('partyAAddress', '#9-11, Street 476, Sangkat Tuol Tumpoung I, Phnom Penh, Cambodia')
            contract.party_b_full_name_with_title = party_b_signature_name
            contract.party_b_address = request.form.get('partyBAddress')
            contract.party_b_phone = request.form.get('partyBPhone')
            contract.party_b_email = request.form.get('partyBEmail')
            contract.registration_number = request.form.get('registrationNumber', '#304 សជណ')
            contract.registration_date = request.form.get('registrationDate', '07 March 2012')
            contract.agreement_start_date = start_date
            contract.agreement_end_date = end_date
            contract.total_fee_usd = total_fee_usd
            contract.gross_amount_usd = total_fee_usd
            contract.tax_percentage = tax_percentage
            contract.payment_installment_desc = '; '.join(payment_installments) if payment_installments else None
            contract.payment_gross = f"USD {total_fee_usd:.2f}" if total_fee_usd else None
            contract.payment_net = f"USD {total_fee_usd * (1 - tax_percentage / 100):.2f}" if total_fee_usd and tax_percentage is not None else None
            contract.workshop_description = request.form.get('workshopDescription')
            contract.focal_person_a_name = request.form.get('focalPersonAName')
            contract.focal_person_a_position = request.form.get('focalPersonAPosition')
            contract.focal_person_a_phone = request.form.get('focalPersonAPhone')
            contract.focal_person_a_email = request.form.get('focalPersonAEmail')
            contract.party_a_signature_name = request.form.get('partyASignatureName', 'Mr. SOEUNG Saroeun')
            contract.party_b_signature_name = party_b_signature_name
            contract.party_b_position = request.form.get('partyBPosition')
            contract.total_fee_words = request.form.get('totalFeeWords')
            contract.title = request.form.get('title')
            contract.deliverables = '; '.join(deliverables) if deliverables else None
            contract.output_description = request.form.get('outputDescription')
            contract.custom_article_sentences = json.dumps({article['articleNumber']: article['customSentence'] for article in articles}) if articles else '{}'

            db.session.commit()
            flash('Contract updated successfully!', 'success')
            return redirect(url_for('contracts.index'))
        except Exception as e:
            db.session.rollback()
            flash(f'Error updating contract: {str(e)}', 'danger')
            return redirect(url_for('contracts.update', contract_id=contract_id))

    contract_data = contract.to_dict()
    contract_data['agreement_start_date_display'] = format_date(contract_data['agreement_start_date'])
    contract_data['agreement_end_date_display'] = format_date(contract_data['agreement_end_date'])
    contract_data['articles'] = [{'articleNumber': k, 'customSentence': v} for k, v in contract_data['custom_article_sentences'].items()]
    return render_template('contracts/edit.html', contract=contract_data)

# Delete contract
@contracts_bp.route('/delete/<string:contract_id>', methods=['POST'])
@login_required
def delete(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    try:
        db.session.delete(contract)
        db.session.commit()
        flash('Contract deleted successfully!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error deleting contract: {str(e)}', 'danger')
    return redirect(url_for('contracts.index'))

# View contract
@contracts_bp.route('/view/<string:contract_id>')
@login_required
def view(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    contract_data = contract.to_dict()
    contract_data['agreement_start_date_display'] = format_date(contract_data['agreement_start_date'])
    contract_data['agreement_end_date_display'] = format_date(contract_data['agreement_end_date'])
    contract_data['articles'] = [{'articleNumber': k, 'customSentence': v} for k, v in contract_data['custom_article_sentences'].items()]
    return render_template('contracts/view.html', contract=contract_data)

# Export contracts to Excel/CSV
@contracts_bp.route('/export/<string:format>')
@login_required
def export(format):
    try:
        search_query = request.args.get('search', '', type=str)
        query = Contract.query
        if search_query:
            query = query.filter(Contract.project_title.ilike(f'%{search_query}%'))

        contracts = [contract.to_dict() for contract in query.all()]
        for contract in contracts:
            contract['agreement_start_date'] = format_date(contract['agreement_start_date'])
            contract['agreement_end_date'] = format_date(contract['agreement_end_date'])
            contract['articles'] = '; '.join([f"Article {a['articleNumber']}: {a['customSentence']}" for a in contract['articles']]) if contract['articles'] else 'N/A'

        data = [{
            'ID': c['id'],
            'Contract Number': c['contract_number'],
            'Project Title': c['project_title'],
            'Output Description': c['output_description'],
            'Tax Percentage (%)': c['tax_percentage'],
            'Party B Signature Name': c['party_b_signature_name'],
            'Party B Position': c['party_b_position'],
            'Party B Phone': c['party_b_phone'],
            'Party B Email': c['party_b_email'],
            'Party B Address': c['party_b_address'],
            'Focal Person A Name': c['focal_person_a_name'],
            'Focal Person A Position': c['focal_person_a_position'],
            'Focal Person A Phone': c['focal_person_a_phone'],
            'Focal Person A Email': c['focal_person_a_email'],
            'Agreement Start Date': c['agreement_start_date'],
            'Agreement End Date': c['agreement_end_date'],
            'Total Fee (USD)': c['total_fee_usd'],
            'Payment Installments': c['payment_installment_desc'],
            'Workshop Description': c['workshop_description'],
            'Deliverables': '; '.join(c['deliverables']) if c['deliverables'] else 'N/A',
            'Articles': c['articles'],
            'Title': c['title']
        } for c in contracts]

        if not data:
            flash('No data available to export.', 'danger')
            return redirect(url_for('contracts.index'))

        if format == 'excel':
            df = pd.DataFrame(data)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            output.seek(0)
            return send_file(output, download_name='contracts.xlsx', as_attachment=True)
        elif format == 'csv':
            df = pd.DataFrame(data)
            output = BytesIO()
            df.to_csv(output, index=False)
            output.seek(0)
            return send_file(output, download_name='contracts.csv', as_attachment=True)
        else:
            flash('Invalid export format.', 'danger')
            return redirect(url_for('contracts.index'))
    except Exception as e:
        flash(f'Export failed: {str(e)}', 'danger')
        return redirect(url_for('contracts.index'))