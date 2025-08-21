from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file
from flask_login import login_required
from app import db
from app.models.contract import Contract
import uuid
import json
from datetime import datetime
import pandas as pd
from io import BytesIO
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

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
    form_data = {}
    if request.method == 'POST':
        try:
            # Collect all form data for re-rendering on error
            form_data = {
                'projectTitle': request.form.get('projectTitle', ''),
                'contractNumber': request.form.get('contractNumber', ''),
                'outputDescription': request.form.get('outputDescription', ''),
                'taxPercentage': request.form.get('taxPercentage', ''),
                'partyBSignatureName': request.form.get('partyBSignatureName', ''),
                'partyBPosition': request.form.get('partyBPosition', ''),
                'partyBPhone': request.form.get('partyBPhone', ''),
                'partyBEmail': request.form.get('partyBEmail', ''),
                'partyBAddress': request.form.get('partyBAddress', ''),
                'focalPersonAName': request.form.get('focalPersonAName', ''),
                'focalPersonAPosition': request.form.get('focalPersonAPosition', ''),
                'focalPersonAPhone': request.form.get('focalPersonAPhone', ''),
                'focalPersonAEmail': request.form.get('focalPersonAEmail', ''),
                'agreementStartDate': request.form.get('agreementStartDate', ''),
                'agreementEndDate': request.form.get('agreementEndDate', ''),
                'totalFeeUSD': request.form.get('totalFeeUSD', ''),
                'totalFeeWords': request.form.get('totalFeeWords', ''),
                'paymentInstallmentDesc': request.form.getlist('paymentInstallmentDesc[]'),
                'workshopDescription': request.form.get('workshopDescription', ''),
                'deliverables': request.form.get('deliverables', ''),
                'articles': [{'articleNumber': num, 'customSentence': sent} for num, sent in zip(request.form.getlist('articleNumber[]'), request.form.getlist('customSentence[]')) if sent.strip()],
                'partyBSignatureNameConfirm': request.form.get('partyBSignatureNameConfirm', ''),
                'title': request.form.get('title', '')
            }

            # Required field validation
            if not form_data['projectTitle']:
                flash('Project title is required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)
            if not form_data['contractNumber']:
                flash('Contract number is required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)
            if not form_data['partyBSignatureName']:
                flash('Party B signature name is required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)
            if not form_data['outputDescription']:
                flash('Output description is required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)
            if not form_data['agreementStartDate']:
                flash('Agreement start date is required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)
            if not form_data['agreementEndDate']:
                flash('Agreement end date is required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)
            if not form_data['totalFeeUSD']:
                flash('Total fee USD is required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)
            if not form_data['deliverables']:
                flash('Deliverables are required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)
            if not form_data['partyBSignatureNameConfirm']:
                flash('Party B signature name confirmation is required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            # Signature confirmation validation
            if form_data['partyBSignatureName'] != form_data['partyBSignatureNameConfirm']:
                flash('Party B Signature Name and Confirmation do not match.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            # Unique contract number validation
            if Contract.query.filter_by(contract_number=form_data['contractNumber']).first():
                flash('Contract number already exists!', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            # Date validation
            start_date = form_data['agreementStartDate']
            end_date = form_data['agreementEndDate']
            if start_date and end_date and datetime.strptime(end_date, '%Y-%m-%d') < datetime.strptime(start_date, '%Y-%m-%d'):
                flash('End date cannot be before start date.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            # Total fee validation
            total_fee_usd = None
            try:
                total_fee_usd = float(form_data['totalFeeUSD']) if form_data['totalFeeUSD'] else None
                if total_fee_usd is not None and total_fee_usd < 0:
                    flash('Total fee cannot be negative.', 'danger')
                    return render_template('contracts/create.html', form_data=form_data)
            except ValueError:
                flash('Invalid total fee amount.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            # Payment installments validation
            payment_installments = [desc.strip() for desc in form_data['paymentInstallmentDesc'] if desc.strip()]
            total_percentage = 0
            for desc in payment_installments:
                try:
                    if '(' in desc and desc.endswith('%)'):
                        percentage = float(desc.split('(')[1].split('%')[0])
                        total_percentage += percentage
                    if total_percentage > 100:
                        flash('Total payment installment percentages cannot exceed 100%.', 'danger')
                        return render_template('contracts/create.html', form_data=form_data)
                except (IndexError, ValueError):
                    flash(f'Invalid format for installment description: {desc}. Please use format like "Installment #1 (50%)".', 'danger')
                    return render_template('contracts/create.html', form_data=form_data)

            # Article validation
            articles = []
            article_numbers = set()
            for article in form_data['articles']:
                try:
                    article_num = int(article['articleNumber'])
                    if article['customSentence'] and article_num not in article_numbers:
                        articles.append(article)
                        article_numbers.add(article_num)
                    elif article['customSentence'] and article_num in article_numbers:
                        flash(f'Duplicate article number {article_num} detected.', 'danger')
                        return render_template('contracts/create.html', form_data=form_data)
                except ValueError:
                    flash(f'Invalid article number: {article["articleNumber"]}.', 'danger')
                    return render_template('contracts/create.html', form_data=form_data)

            # Deliverables processing
            deliverables = form_data['deliverables'].split('\n')
            deliverables = [d.strip() for d in deliverables if d.strip()]
            if not deliverables:
                flash('At least one deliverable is required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            # Numeric fields
            tax_percentage = float(form_data['taxPercentage']) if form_data['taxPercentage'] else None

            # Create new contract
            new_contract = Contract(
                id=str(uuid.uuid4()),
                project_title=form_data['projectTitle'],
                contract_number=form_data['contractNumber'],
                organization_name=request.form.get('organizationName', 'The NGO Forum on Cambodia'),
                party_a_name=request.form.get('partyAName', 'Mr. Soeung Saroeun'),
                party_a_position=request.form.get('partyAPosition', 'Executive Director'),
                party_a_address=request.form.get('partyAAddress', '#9-11, Street 476, Sangkat Tuol Tumpoung I, Phnom Penh, Cambodia'),
                party_b_full_name_with_title=form_data['partyBSignatureName'],
                party_b_address=form_data['partyBAddress'],
                party_b_phone=form_data['partyBPhone'],
                party_b_email=form_data['partyBEmail'],
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
                workshop_description=form_data['workshopDescription'],
                focal_person_a_name=form_data['focalPersonAName'],
                focal_person_a_position=form_data['focalPersonAPosition'],
                focal_person_a_phone=form_data['focalPersonAPhone'],
                focal_person_a_email=form_data['focalPersonAEmail'],
                party_a_signature_name=request.form.get('partyASignatureName', 'Mr. SOEUNG Saroeun'),
                party_b_signature_name=form_data['partyBSignatureName'],
                party_b_position=form_data['partyBPosition'],
                total_fee_words=form_data['totalFeeWords'],
                title=form_data['title'],
                deliverables='; '.join(deliverables) if deliverables else None,
                output_description=form_data['outputDescription'],
                custom_article_sentences=json.dumps({article['articleNumber']: article['customSentence'] for article in articles}) if articles else '{}'
            )

            db.session.add(new_contract)
            db.session.commit()
            flash('Contract created successfully!', 'success')
            return redirect(url_for('contracts.index'))
        except Exception as e:
            db.session.rollback()
            logger.error(f'Error creating contract: {str(e)}')
            flash(f'Error creating contract: {str(e)}', 'danger')
            return render_template('contracts/create.html', form_data=form_data)

    return render_template('contracts/create.html', form_data=form_data)

# Update contract
@contracts_bp.route('/update/<string:contract_id>', methods=['GET', 'POST'])
@login_required
def update(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    form_data = {}
    if request.method == 'POST':
        try:
            # Collect all form data for re-rendering on error
            form_data = {
                'id': contract_id,
                'projectTitle': request.form.get('projectTitle', ''),
                'contractNumber': request.form.get('contractNumber', ''),
                'outputDescription': request.form.get('outputDescription', ''),
                'taxPercentage': request.form.get('taxPercentage', ''),
                'partyBSignatureName': request.form.get('partyBSignatureName', ''),
                'partyBPosition': request.form.get('partyBPosition', ''),
                'partyBPhone': request.form.get('partyBPhone', ''),
                'partyBEmail': request.form.get('partyBEmail', ''),
                'partyBAddress': request.form.get('partyBAddress', ''),
                'focalPersonAName': request.form.get('focalPersonAName', ''),
                'focalPersonAPosition': request.form.get('focalPersonAPosition', ''),
                'focalPersonAPhone': request.form.get('focalPersonAPhone', ''),
                'focalPersonAEmail': request.form.get('focalPersonAEmail', ''),
                'agreementStartDate': request.form.get('agreementStartDate', ''),
                'agreementEndDate': request.form.get('agreementEndDate', ''),
                'totalFeeUSD': request.form.get('totalFeeUSD', ''),
                'totalFeeWords': request.form.get('totalFeeWords', ''),
                'paymentInstallmentDesc': request.form.getlist('paymentInstallmentDesc[]'),
                'workshopDescription': request.form.get('workshopDescription', ''),
                'deliverables': request.form.get('deliverables', ''),
                'articles': [{'articleNumber': num, 'customSentence': sent} for num, sent in zip(request.form.getlist('articleNumber[]'), request.form.getlist('customSentence[]')) if sent.strip()],
                'partyBSignatureNameConfirm': request.form.get('partyBSignatureNameConfirm', ''),
                'title': request.form.get('title', '')
            }

            # Required field validation
            if not form_data['projectTitle']:
                flash('Project title is required.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)
            if not form_data['contractNumber']:
                flash('Contract number is required.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)
            if not form_data['partyBSignatureName']:
                flash('Party B signature name is required.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)
            if not form_data['outputDescription']:
                flash('Output description is required.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)
            if not form_data['agreementStartDate']:
                flash('Agreement start date is required.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)
            if not form_data['agreementEndDate']:
                flash('Agreement end date is required.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)
            if not form_data['totalFeeUSD']:
                flash('Total fee USD is required.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)
            if not form_data['deliverables']:
                flash('Deliverables are required.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)
            if not form_data['partyBSignatureNameConfirm']:
                flash('Party B signature name confirmation is required.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)

            # Signature confirmation validation
            if form_data['partyBSignatureName'] != form_data['partyBSignatureNameConfirm']:
                flash('Party B Signature Name and Confirmation do not match.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)

            # Unique contract number validation
            existing_contract = Contract.query.filter_by(contract_number=form_data['contractNumber']).first()
            if existing_contract and existing_contract.id != contract_id:
                flash('Contract number already exists!', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)

            # Date validation
            start_date = form_data['agreementStartDate']
            end_date = form_data['agreementEndDate']
            if start_date and end_date:
                try:
                    if datetime.strptime(end_date, '%Y-%m-%d') < datetime.strptime(start_date, '%Y-%m-%d'):
                        flash('End date cannot be before start date.', 'danger')
                        return render_template('contracts/edit.html', form_data=form_data)
                except ValueError as e:
                    logger.error(f'Date parsing error: {str(e)}')
                    flash('Invalid date format.', 'danger')
                    return render_template('contracts/edit.html', form_data=form_data)

            # Total fee validation
            total_fee_usd = None
            try:
                total_fee_usd = float(form_data['totalFeeUSD']) if form_data['totalFeeUSD'] else None
                if total_fee_usd is not None and total_fee_usd < 0:
                    flash('Total fee cannot be negative.', 'danger')
                    return render_template('contracts/edit.html', form_data=form_data)
            except ValueError:
                flash('Invalid total fee amount.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)

            # Payment installments validation
            payment_installments = [desc.strip() for desc in form_data['paymentInstallmentDesc'] if desc.strip()]
            total_percentage = 0
            for desc in payment_installments:
                try:
                    if '(' in desc and desc.endswith('%)'):
                        percentage = float(desc.split('(')[1].split('%')[0])
                        total_percentage += percentage
                    if total_percentage > 100:
                        flash('Total payment installment percentages cannot exceed 100%.', 'danger')
                        return render_template('contracts/edit.html', form_data=form_data)
                except (IndexError, ValueError):
                    flash(f'Invalid format for installment description: {desc}. Please use format like "Installment #1 (50%)".', 'danger')
                    return render_template('contracts/edit.html', form_data=form_data)

            # Article validation
            articles = []
            article_numbers = set()
            for article in form_data['articles']:
                try:
                    article_num = int(article['articleNumber'])
                    if article['customSentence'] and article_num not in article_numbers:
                        articles.append(article)
                        article_numbers.add(article_num)
                    elif article['customSentence'] and article_num in article_numbers:
                        flash(f'Duplicate article number {article_num} detected.', 'danger')
                        return render_template('contracts/edit.html', form_data=form_data)
                except ValueError:
                    flash(f'Invalid article number: {article["articleNumber"]}.', 'danger')
                    return render_template('contracts/edit.html', form_data=form_data)

            # Deliverables processing
            deliverables = form_data['deliverables'].split('\n')
            deliverables = [d.strip() for d in deliverables if d.strip()]
            if not deliverables:
                flash('At least one deliverable is required.', 'danger')
                return render_template('contracts/edit.html', form_data=form_data)

            # Numeric fields
            tax_percentage = float(form_data['taxPercentage']) if form_data['taxPercentage'] else None

            # Update contract
            contract.project_title = form_data['projectTitle']
            contract.contract_number = form_data['contractNumber']
            contract.organization_name = request.form.get('organizationName', contract.organization_name or 'The NGO Forum on Cambodia')
            contract.party_a_name = request.form.get('partyAName', contract.party_a_name or 'Mr. Soeung Saroeun')
            contract.party_a_position = request.form.get('partyAPosition', contract.party_a_position or 'Executive Director')
            contract.party_a_address = request.form.get('partyAAddress', contract.party_a_address or '#9-11, Street 476, Sangkat Tuol Tumpoung I, Phnom Penh, Cambodia')
            contract.party_b_full_name_with_title = form_data['partyBSignatureName']
            contract.party_b_address = form_data['partyBAddress']
            contract.party_b_phone = form_data['partyBPhone']
            contract.party_b_email = form_data['partyBEmail']
            contract.registration_number = request.form.get('registrationNumber', contract.registration_number or '#304 សជណ')
            contract.registration_date = request.form.get('registrationDate', contract.registration_date or '07 March 2012')
            contract.agreement_start_date = start_date
            contract.agreement_end_date = end_date
            contract.total_fee_usd = total_fee_usd
            contract.gross_amount_usd = total_fee_usd
            contract.tax_percentage = tax_percentage
            contract.payment_installment_desc = '; '.join(payment_installments) if payment_installments else None
            contract.payment_gross = f"USD {total_fee_usd:.2f}" if total_fee_usd else None
            contract.payment_net = f"USD {total_fee_usd * (1 - tax_percentage / 100):.2f}" if total_fee_usd and tax_percentage is not None else None
            contract.workshop_description = form_data['workshopDescription']
            contract.focal_person_a_name = form_data['focalPersonAName']
            contract.focal_person_a_position = form_data['focalPersonAPosition']
            contract.focal_person_a_phone = form_data['focalPersonAPhone']
            contract.focal_person_a_email = form_data['focalPersonAEmail']
            contract.party_a_signature_name = request.form.get('partyASignatureName', contract.party_a_signature_name or 'Mr. SOEUNG Saroeun')
            contract.party_b_signature_name = form_data['partyBSignatureName']
            contract.party_b_position = form_data['partyBPosition']
            contract.total_fee_words = form_data['totalFeeWords']
            contract.title = form_data['title']
            contract.deliverables = '; '.join(deliverables) if deliverables else None
            contract.output_description = form_data['outputDescription']
            contract.custom_article_sentences = json.dumps({article['articleNumber']: article['customSentence'] for article in articles}) if articles else '{}'

            db.session.commit()
            flash('Contract updated successfully!', 'success')
            return redirect(url_for('contracts.index'))
        except Exception as e:
            db.session.rollback()
            logger.error(f'Error updating contract {contract_id}: {str(e)}')
            flash(f'Error updating contract: {str(e)}', 'danger')
            return render_template('contracts/edit.html', form_data=form_data)

    # Prepare form_data for GET request
    contract_data = contract.to_dict()
    contract_data['agreement_start_date_display'] = format_date(contract_data['agreement_start_date'])
    contract_data['agreement_end_date_display'] = format_date(contract_data['agreement_end_date'])
    contract_data['articles'] = [{'articleNumber': k, 'customSentence': v} for k, v in contract_data['custom_article_sentences'].items()] if contract_data['custom_article_sentences'] else []
    contract_data['paymentInstallmentDesc'] = contract_data['payment_installment_desc'].split('; ') if contract_data['payment_installment_desc'] else ['']
    # Handle deliverables: expect a string, but account for possible list from to_dict
    if isinstance(contract_data['deliverables'], list):
        contract_data['deliverables'] = '\n'.join(contract_data['deliverables']) if contract_data['deliverables'] else ''
    else:
        contract_data['deliverables'] = '\n'.join(contract_data['deliverables'].split('; ')) if contract_data['deliverables'] else ''
    contract_data['totalFeeUSD'] = contract_data['total_fee_usd']
    contract_data['totalFeeWords'] = contract_data['total_fee_words'] or ''
    contract_data['projectTitle'] = contract_data['project_title']
    contract_data['contractNumber'] = contract_data['contract_number']
    contract_data['outputDescription'] = contract_data['output_description']
    contract_data['taxPercentage'] = str(contract_data['tax_percentage']) if contract_data['tax_percentage'] is not None else ''
    contract_data['partyBSignatureName'] = contract_data['party_b_signature_name']
    contract_data['partyBPosition'] = contract_data['party_b_position']
    contract_data['partyBPhone'] = contract_data['party_b_phone']
    contract_data['partyBEmail'] = contract_data['party_b_email']
    contract_data['partyBAddress'] = contract_data['party_b_address']
    contract_data['focalPersonAName'] = contract_data['focal_person_a_name']
    contract_data['focalPersonAPosition'] = contract_data['focal_person_a_position']
    contract_data['focalPersonAPhone'] = contract_data['focal_person_a_phone']
    contract_data['focalPersonAEmail'] = contract_data['focal_person_a_email']
    contract_data['agreementStartDate'] = contract_data['agreement_start_date']
    contract_data['agreementEndDate'] = contract_data['agreement_end_date']
    contract_data['workshopDescription'] = contract_data['workshop_description']
    contract_data['partyBSignatureNameConfirm'] = contract_data['party_b_signature_name']
    contract_data['title'] = contract_data['title']
    logger.debug(f"form_data for GET: {contract_data}")
    return render_template('contracts/edit.html', form_data=contract_data)

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
        logger.error(f'Error deleting contract {contract_id}: {str(e)}')
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
            'Total Fee in Words': c['total_fee_words'],
            'Payment Installments': c['payment_installment_desc'],
            'Workshop Description': c['workshop_description'],
            'Deliverables': c['deliverables'],
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
        logger.error(f'Error exporting contracts: {str(e)}')
        flash(f'Export failed: {str(e)}', 'danger')
        return redirect(url_for('contracts.index'))