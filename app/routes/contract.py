from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file
from flask_login import login_required
from app import db
from app.models.contract import Contract
import uuid
from datetime import datetime
import pandas as pd
from io import BytesIO
import logging
from num2words import num2words
import re

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

# Helper function to convert number to words
def number_to_words(num):
    try:
        if not num or num < 0:
            return "Zero US Dollars only"
        integer_part = int(num)
        decimal_part = round((num - integer_part) * 100)
        words = num2words(integer_part, lang='en').title()
        if decimal_part > 0:
            words += " and " + num2words(decimal_part, lang='en').title() + " Cents"
        return f"{words} US Dollars only"
    except Exception as e:
        logger.error(f"Error converting number to words: {str(e)}")
        return "N/A"

# Helper function to normalize field to list
def normalize_to_list(field):
    if isinstance(field, list):
        return [str(item).strip() for item in field if str(item).strip()]
    elif isinstance(field, str):
        return [item.strip() for item in field.split('; ') if item.strip()]
    return []

# Helper function to calculate payment gross and net
def calculate_payments(total_fee_usd, tax_percentage, payment_installments):
    try:
        total_gross = 0.0
        total_net = 0.0
        for installment in payment_installments:
            match = re.search(r'\((\d+\.?\d*)\%\)', installment['description'])
            if not match:
                raise ValueError(f"Invalid percentage format in installment: {installment['description']}")
            percentage = float(match.group(1))
            gross_amount = (total_fee_usd * percentage) / 100
            net_amount = gross_amount * (1 - tax_percentage / 100)
            total_gross += gross_amount
            total_net += net_amount
        return f"${total_gross:.2f} USD", f"${total_net:.2f} USD"
    except Exception as e:
        logger.error(f"Error calculating payments: {str(e)}")
        raise

# List contracts with pagination, search, and sorting
@contracts_bp.route('/')
@login_required
def index():
    try:
        page = request.args.get('page', 1, type=int)
        search_query = request.args.get('search', '', type=str)
        sort_order = request.args.get('sort', 'project_title_asc', type=str)
        entries_per_page = request.args.get('entries', 10, type=int)

        query = Contract.query
        if search_query:
            query = query.filter(
                (Contract.project_title.ilike(f'%{search_query}%')) |
                (Contract.contract_number.ilike(f'%{search_query}%')) |
                (Contract.party_b_signature_name.ilike(f'%{search_query}%'))
            )

        if sort_order == 'project_title_asc':
            query = query.order_by(Contract.project_title.asc())
        elif sort_order == 'project_title_desc':
            query = query.order_by(Contract.project_title.desc())
        elif sort_order == 'start_date_asc':
            query = query.order_by(Contract.agreement_start_date.asc())
        elif sort_order == 'start_date_desc':
            query = query.order_by(Contract.agreement_start_date.desc())
        elif sort_order == 'total_fee_asc':
            query = query.order_by(Contract.total_fee_usd.asc())
        elif sort_order == 'total_fee_desc':
            query = query.order_by(Contract.total_fee_usd.desc())

        pagination = query.paginate(page=page, per_page=entries_per_page, error_out=False)
        contracts = [contract.to_dict() for contract in pagination.items]
        for contract in contracts:
            contract['agreement_start_date_display'] = format_date(contract['agreement_start_date'])
            contract['agreement_end_date_display'] = format_date(contract['agreement_end_date'])
            contract['total_fee_usd'] = f"{contract['total_fee_usd']:.2f}" if contract['total_fee_usd'] is not None else '0.00'

        return render_template('contracts/index.html', contracts=contracts, pagination=pagination,
                              search_query=search_query, sort_order=sort_order, entries_per_page=entries_per_page)
    except Exception as e:
        logger.error(f"Error in index route: {str(e)}")
        flash(f"Error loading contracts: {str(e)}", 'danger')
        return render_template('contracts/index.html', contracts=[], pagination=None,
                              search_query='', sort_order='project_title_asc', entries_per_page=10)

# Create contract
@contracts_bp.route('/create', methods=['GET', 'POST'])
@login_required
def create():
    form_data = {}
    if request.method == 'POST':
        try:
            form_data = {
                'project_title': request.form.get('project_title', '').strip(),
                'contract_number': request.form.get('contract_number', '').strip(),
                'output_description': request.form.get('output_description', '').strip(),
                'tax_percentage': request.form.get('tax_percentage', '').strip(),
                'organization_name': request.form.get('organization_name', '').strip(),
                'party_a_name': request.form.get('party_a_name', '').strip(),
                'party_a_position': request.form.get('party_a_position', '').strip(),
                'party_a_address': request.form.get('party_a_address', '').strip(),
                'party_b_signature_name': request.form.get('party_b_signature_name', '').strip(),
                'party_b_position': request.form.get('party_b_position', '').strip(),
                'party_b_phone': request.form.get('party_b_phone', '').strip(),
                'party_b_email': request.form.get('party_b_email', '').strip(),
                'party_b_address': request.form.get('party_b_address', '').strip(),
                'focal_person_a_name': request.form.get('focal_person_a_name', '').strip(),
                'focal_person_a_position': request.form.get('focal_person_a_position', '').strip(),
                'focal_person_a_phone': request.form.get('focal_person_a_phone', '').strip(),
                'focal_person_a_email': request.form.get('focal_person_a_email', '').strip(),
                'agreement_start_date': request.form.get('agreement_start_date', '').strip(),
                'agreement_end_date': request.form.get('agreement_end_date', '').strip(),
                'total_fee_usd': request.form.get('total_fee_usd', '').strip(),
                'total_fee_words': request.form.get('total_fee_words', '').strip(),
                'payment_installments': [
                    {
                        'description': desc.strip(),
                        'deliverables': deliv.strip(),
                        'dueDate': due.strip()
                    }
                    for desc, deliv, due in zip(
                        request.form.getlist('paymentInstallmentDesc[]'),
                        request.form.getlist('paymentInstallmentDeliverables[]'),
                        request.form.getlist('paymentInstallmentDueDate[]')
                    )
                    if desc.strip() and deliv.strip() and due.strip()
                ],
                'workshop_description': request.form.get('workshop_description', '').strip(),
                'articles': [
                    {'article_number': num.strip(), 'custom_sentence': sent.strip()}
                    for num, sent in zip(request.form.getlist('articleNumber[]'), request.form.getlist('customSentence[]'))
                    if sent.strip()
                ],
                'party_b_signature_name_confirm': request.form.get('party_b_signature_name_confirm', '').strip(),
                'title': request.form.get('title', '').strip()
            }

            required_fields = [
                ('project_title', 'Project title is required.'),
                ('contract_number', 'Contract number is required.'),
                ('output_description', 'Output description is required.'),
                ('organization_name', 'Organization name is required.'),
                ('party_a_name', 'Party A name is required.'),
                ('party_a_position', 'Party A position is required.'),
                ('party_a_address', 'Party A address is required.'),
                ('party_b_signature_name', 'Party B signature name is required.'),
                ('agreement_start_date', 'Agreement start date is required.'),
                ('agreement_end_date', 'Agreement end date is required.'),
                ('total_fee_usd', 'Total fee USD is required.'),
                ('party_b_signature_name_confirm', 'Party B signature name confirmation is required.')
            ]
            for field, message in required_fields:
                if not form_data[field]:
                    flash(message, 'danger')
                    return render_template('contracts/create.html', form_data=form_data)

            if not form_data['payment_installments']:
                flash('At least one payment installment is required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            if form_data['party_b_signature_name'] != form_data['party_b_signature_name_confirm']:
                flash('Party B Signature Name and Confirmation do not match.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            if Contract.query.filter_by(contract_number=form_data['contract_number']).first():
                flash('Contract number already exists.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            start_date = form_data['agreement_start_date']
            end_date = form_data['agreement_end_date']
            if start_date and end_date:
                try:
                    if datetime.strptime(end_date, '%Y-%m-%d') < datetime.strptime(start_date, '%Y-%m-%d'):
                        flash('End date cannot be before start date.', 'danger')
                        return render_template('contracts/create.html', form_data=form_data)
                except ValueError:
                    flash('Invalid date format.', 'danger')
                    return render_template('contracts/create.html', form_data=form_data)

            try:
                total_fee_usd = float(form_data['total_fee_usd'])
                if total_fee_usd < 0:
                    flash('Total fee cannot be negative.', 'danger')
                    return render_template('contracts/create.html', form_data=form_data)
            except ValueError:
                flash('Invalid total fee amount.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            try:
                tax_percentage = float(form_data['tax_percentage']) if form_data['tax_percentage'] else 15.0
                if tax_percentage < 0:
                    flash('Tax percentage cannot be negative.', 'danger')
                    return render_template('contracts/create.html', form_data=form_data)
            except ValueError:
                flash('Invalid tax percentage.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            total_percentage = 0
            for installment in form_data['payment_installments']:
                try:
                    match = re.search(r'\((\d+\.?\d*)\%\)', installment['description'])
                    if not match:
                        flash(f'Invalid format for installment description: {installment["description"]}. Please use format like "Installment #1 (50%)".', 'danger')
                        return render_template('contracts/create.html', form_data=form_data)
                    percentage = float(match.group(1))
                    total_percentage += percentage
                    if not installment['deliverables']:
                        flash('Each installment must have deliverables.', 'danger')
                        return render_template('contracts/create.html', form_data=form_data)
                    if not installment['dueDate']:
                        flash('Each installment must have a due date.', 'danger')
                        return render_template('contracts/create.html', form_data=form_data)
                    try:
                        datetime.strptime(installment['dueDate'], '%Y-%m-%d')
                    except ValueError:
                        flash(f'Invalid due date format for installment: {installment["description"]}.', 'danger')
                        return render_template('contracts/create.html', form_data=form_data)
                except (IndexError, ValueError):
                    flash(f'Invalid format for installment description: {installment["description"]}. Please use format like "Installment #1 (50%)".', 'danger')
                    return render_template('contracts/create.html', form_data=form_data)

            if abs(total_percentage - 100) > 0.01:
                flash(f'Total percentage of installments must equal 100%, but got {total_percentage}%.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            payment_gross, payment_net = calculate_payments(total_fee_usd, tax_percentage, form_data['payment_installments'])

            try:
                custom_article_sentences = {str(article['article_number']): article['custom_sentence'] for article in form_data['articles']}
                contract = Contract(
                    id=str(uuid.uuid4()),
                    project_title=form_data['project_title'],
                    contract_number=form_data['contract_number'],
                    output_description=form_data['output_description'],
                    tax_percentage=tax_percentage,
                    organization_name=form_data['organization_name'],
                    party_a_name=form_data['party_a_name'],
                    party_a_position=form_data['party_a_position'],
                    party_a_address=form_data['party_a_address'],
                    party_b_signature_name=form_data['party_b_signature_name'],
                    party_b_position=form_data['party_b_position'],
                    party_b_phone=form_data['party_b_phone'],
                    party_b_email=form_data['party_b_email'],
                    party_b_address=form_data['party_b_address'],
                    focal_person_a_name=form_data['focal_person_a_name'],
                    focal_person_a_position=form_data['focal_person_a_position'],
                    focal_person_a_phone=form_data['focal_person_a_phone'],
                    focal_person_a_email=form_data['focal_person_a_email'],
                    agreement_start_date=form_data['agreement_start_date'],
                    agreement_end_date=form_data['agreement_end_date'],
                    total_fee_usd=total_fee_usd,
                    gross_amount_usd=total_fee_usd,
                    total_fee_words=number_to_words(total_fee_usd),
                    payment_installments=form_data['payment_installments'],
                    payment_gross=payment_gross,
                    payment_net=payment_net,
                    workshop_description=form_data['workshop_description'],
                    title=form_data['title'],
                    custom_article_sentences=custom_article_sentences,
                    deliverables='; '.join(normalize_to_list(form_data.get('deliverables', '')))
                )
                db.session.add(contract)
                db.session.commit()
                flash('Contract created successfully!', 'success')
                return redirect(url_for('contracts.index'))
            except Exception as e:
                db.session.rollback()
                logger.error(f"Error saving contract: {str(e)}")
                flash(f"Error saving contract: {str(e)}", 'danger')
                return render_template('contracts/create.html', form_data=form_data)

        except Exception as e:
            logger.error(f"Error in create route: {str(e)}")
            flash(f"Error processing form: {str(e)}", 'danger')
            return render_template('contracts/create.html', form_data=form_data)

    return render_template('contracts/create.html', form_data=form_data)

# Update contract
@contracts_bp.route('/edit/<contract_id>', methods=['GET', 'POST'])
@login_required
def update(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    form_data = contract.to_dict()

    if request.method == 'POST':
        try:
            form_data = {
                'id': contract_id,
                'project_title': request.form.get('project_title', '').strip(),
                'contract_number': request.form.get('contract_number', '').strip(),
                'output_description': request.form.get('output_description', '').strip(),
                'tax_percentage': request.form.get('tax_percentage', '').strip(),
                'organization_name': request.form.get('organization_name', '').strip(),
                'party_a_name': request.form.get('party_a_name', '').strip(),
                'party_a_position': request.form.get('party_a_position', '').strip(),
                'party_a_address': request.form.get('party_a_address', '').strip(),
                'party_b_signature_name': request.form.get('party_b_signature_name', '').strip(),
                'party_b_position': request.form.get('party_b_position', '').strip(),
                'party_b_phone': request.form.get('party_b_phone', '').strip(),
                'party_b_email': request.form.get('party_b_email', '').strip(),
                'party_b_address': request.form.get('party_b_address', '').strip(),
                'focal_person_a_name': request.form.get('focal_person_a_name', '').strip(),
                'focal_person_a_position': request.form.get('focal_person_a_position', '').strip(),
                'focal_person_a_phone': request.form.get('focal_person_a_phone', '').strip(),
                'focal_person_a_email': request.form.get('focal_person_a_email', '').strip(),
                'agreement_start_date': request.form.get('agreement_start_date', '').strip(),
                'agreement_end_date': request.form.get('agreement_end_date', '').strip(),
                'total_fee_usd': request.form.get('total_fee_usd', '').strip(),
                'total_fee_words': request.form.get('total_fee_words', '').strip(),
                'payment_installments': [
                    {
                        'description': desc.strip(),
                        'deliverables': deliv.strip(),
                        'dueDate': due.strip()
                    }
                    for desc, deliv, due in zip(
                        request.form.getlist('paymentInstallmentDesc[]'),
                        request.form.getlist('paymentInstallmentDeliverables[]'),
                        request.form.getlist('paymentInstallmentDueDate[]')
                    )
                    if desc.strip() and deliv.strip() and due.strip()
                ],
                'workshop_description': request.form.get('workshop_description', '').strip(),
                'articles': [
                    {'article_number': num.strip(), 'custom_sentence': sent.strip()}
                    for num, sent in zip(request.form.getlist('articleNumber[]'), request.form.getlist('customSentence[]'))
                    if sent.strip()
                ],
                'party_b_signature_name_confirm': request.form.get('party_b_signature_name_confirm', '').strip(),
                'title': request.form.get('title', '').strip()
            }

            required_fields = [
                ('project_title', 'Project title is required.'),
                ('contract_number', 'Contract number is required.'),
                ('output_description', 'Output description is required.'),
                ('organization_name', 'Organization name is required.'),
                ('party_a_name', 'Party A name is required.'),
                ('party_a_position', 'Party A position is required.'),
                ('party_a_address', 'Party A address is required.'),
                ('party_b_signature_name', 'Party B signature name is required.'),
                ('agreement_start_date', 'Agreement start date is required.'),
                ('agreement_end_date', 'Agreement end date is required.'),
                ('total_fee_usd', 'Total fee USD is required.'),
                ('party_b_signature_name_confirm', 'Party B signature name confirmation is required.')
            ]
            for field, message in required_fields:
                if not form_data[field]:
                    flash(message, 'danger')
                    return render_template('contracts/update.html', form_data=form_data)

            if not form_data['payment_installments']:
                flash('At least one payment installment is required.', 'danger')
                return render_template('contracts/update.html', form_data=form_data)

            if form_data['party_b_signature_name'] != form_data['party_b_signature_name_confirm']:
                flash('Party B Signature Name and Confirmation do not match.', 'danger')
                return render_template('contracts/update.html', form_data=form_data)

            existing_contract = Contract.query.filter(Contract.contract_number == form_data['contract_number'], Contract.id != contract_id).first()
            if existing_contract:
                flash('Contract number already exists.', 'danger')
                return render_template('contracts/update.html', form_data=form_data)

            start_date = form_data['agreement_start_date']
            end_date = form_data['agreement_end_date']
            if start_date and end_date:
                try:
                    if datetime.strptime(end_date, '%Y-%m-%d') < datetime.strptime(start_date, '%Y-%m-%d'):
                        flash('End date cannot be before start date.', 'danger')
                        return render_template('contracts/update.html', form_data=form_data)
                except ValueError:
                    flash('Invalid date format.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)

            try:
                total_fee_usd = float(form_data['total_fee_usd'])
                if total_fee_usd < 0:
                    flash('Total fee cannot be negative.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)
            except ValueError:
                flash('Invalid total fee amount.', 'danger')
                return render_template('contracts/update.html', form_data=form_data)

            try:
                tax_percentage = float(form_data['tax_percentage']) if form_data['tax_percentage'] else 15.0
                if tax_percentage < 0:
                    flash('Tax percentage cannot be negative.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)
            except ValueError:
                flash('Invalid tax percentage.', 'danger')
                return render_template('contracts/update.html', form_data=form_data)

            total_percentage = 0
            for installment in form_data['payment_installments']:
                try:
                    match = re.search(r'\((\d+\.?\d*)\%\)', installment['description'])
                    if not match:
                        flash(f'Invalid format for installment description: {installment["description"]}. Please use format like "Installment #1 (50%)".', 'danger')
                        return render_template('contracts/update.html', form_data=form_data)
                    percentage = float(match.group(1))
                    total_percentage += percentage
                    if not installment['deliverables']:
                        flash('Each installment must have deliverables.', 'danger')
                        return render_template('contracts/update.html', form_data=form_data)
                    if not installment['dueDate']:
                        flash('Each installment must have a due date.', 'danger')
                        return render_template('contracts/update.html', form_data=form_data)
                    try:
                        datetime.strptime(installment['dueDate'], '%Y-%m-%d')
                    except ValueError:
                        flash(f'Invalid due date format for installment: {installment["description"]}.', 'danger')
                        return render_template('contracts/update.html', form_data=form_data)
                except (IndexError, ValueError):
                    flash(f'Invalid format for installment description: {installment["description"]}. Please use format like "Installment #1 (50%)".', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)

            if abs(total_percentage - 100) > 0.01:
                flash(f'Total percentage of installments must equal 100%, but got {total_percentage}%.', 'danger')
                return render_template('contracts/update.html', form_data=form_data)

            payment_gross, payment_net = calculate_payments(total_fee_usd, tax_percentage, form_data['payment_installments'])

            try:
                custom_article_sentences = {str(article['article_number']): article['custom_sentence'] for article in form_data['articles']}
                contract.project_title = form_data['project_title']
                contract.contract_number = form_data['contract_number']
                contract.output_description = form_data['output_description']
                contract.tax_percentage = tax_percentage
                contract.organization_name = form_data['organization_name']
                contract.party_a_name = form_data['party_a_name']
                contract.party_a_position = form_data['party_a_position']
                contract.party_a_address = form_data['party_a_address']
                contract.party_b_signature_name = form_data['party_b_signature_name']
                contract.party_b_position = form_data['party_b_position']
                contract.party_b_phone = form_data['party_b_phone']
                contract.party_b_email = form_data['party_b_email']
                contract.party_b_address = form_data['party_b_address']
                contract.focal_person_a_name = form_data['focal_person_a_name']
                contract.focal_person_a_position = form_data['focal_person_a_position']
                contract.focal_person_a_phone = form_data['focal_person_a_phone']
                contract.focal_person_a_email = form_data['focal_person_a_email']
                contract.agreement_start_date = form_data['agreement_start_date']
                contract.agreement_end_date = form_data['agreement_end_date']
                contract.total_fee_usd = total_fee_usd
                contract.gross_amount_usd = total_fee_usd
                contract.total_fee_words = number_to_words(total_fee_usd)
                contract.payment_installments = form_data['payment_installments']
                contract.payment_gross = payment_gross
                contract.payment_net = payment_net
                contract.workshop_description = form_data['workshop_description']
                contract.title = form_data['title']
                contract.custom_article_sentences = custom_article_sentences
                contract.deliverables = '; '.join(normalize_to_list(form_data.get('deliverables', '')))
                db.session.commit()
                flash('Contract updated successfully!', 'success')
                return redirect(url_for('contracts.index'))
            except Exception as e:
                db.session.rollback()
                logger.error(f"Error updating contract: {str(e)}")
                flash(f"Error updating contract: {str(e)}", 'danger')
                return render_template('contracts/update.html', form_data=form_data)

        except Exception as e:
            logger.error(f"Error in update route: {str(e)}")
            flash(f"Error processing form: {str(e)}", 'danger')
            return render_template('contracts/update.html', form_data=form_data)

    return render_template('contracts/update.html', form_data=form_data)

# View contract
@contracts_bp.route('/view/<contract_id>')
@login_required
def view(contract_id):
    try:
        contract = Contract.query.get_or_404(contract_id)
        contract_data = contract.to_dict()
        contract_data['agreement_start_date_display'] = format_date(contract_data['agreement_start_date'])
        contract_data['agreement_end_date_display'] = format_date(contract_data['agreement_end_date'])
        contract_data['total_fee_usd'] = f"{contract_data['total_fee_usd']:.2f}" if contract_data['total_fee_usd'] is not None else '0.00'
        for installment in contract_data.get('payment_installments', []):
            installment['dueDate'] = format_date(installment.get('dueDate', ''))
        return render_template('contracts/view.html', contract=contract_data)
    except Exception as e:
        logger.error(f"Error in view route: {str(e)}")
        flash(f"Error viewing contract: {str(e)}", 'danger')
        return redirect(url_for('contracts.index'))

# Delete contract
@contracts_bp.route('/delete/<contract_id>', methods=['POST'])
@login_required
def delete(contract_id):
    try:
        contract = Contract.query.get_or_404(contract_id)
        db.session.delete(contract)
        db.session.commit()
        flash('Contract deleted successfully!', 'success')
        return redirect(url_for('contracts.index'))
    except Exception as e:
        db.session.rollback()
        logger.error(f"Error deleting contract: {str(e)}")
        flash(f"Error deleting contract: {str(e)}", 'danger')
        return redirect(url_for('contracts.index'))