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
from num2words import num2words
import re
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
import docx

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

# List contracts with pagination, search, and sorting
@contracts_bp.route('/')
@login_required
def index():
    try:
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
            contract['total_fee_usd'] = f"{contract['total_fee_usd']:.2f}" if contract['total_fee_usd'] is not None else '0.00'

        return render_template('contracts/index.html', contracts=contracts, pagination=pagination,
                              search_query=search_query, sort_order=sort_order, entries_per_page=entries_per_page)
    except Exception as e:
        logger.error(f"Error in index route: {str(e)}")
        flash(f"Error loading contracts: {str(e)}", 'danger')
        return render_template('contracts/index.html', contracts=[], pagination=None,
                              search_query='', sort_order='asc', entries_per_page=10)

# Create contract
@contracts_bp.route('/create', methods=['GET', 'POST'])
@login_required
def create():
    form_data = {}
    if request.method == 'POST':
        try:
            # Collect all form data for re-rendering on error
            form_data = {
                'projectTitle': request.form.get('projectTitle', '').strip(),
                'contractNumber': request.form.get('contractNumber', '').strip(),
                'outputDescription': request.form.get('outputDescription', '').strip(),
                'taxPercentage': request.form.get('taxPercentage', '').strip(),
                'partyBSignatureName': request.form.get('partyBSignatureName', '').strip(),
                'partyBPosition': request.form.get('partyBPosition', '').strip(),
                'partyBPhone': request.form.get('partyBPhone', '').strip(),
                'partyBEmail': request.form.get('partyBEmail', '').strip(),
                'partyBAddress': request.form.get('partyBAddress', '').strip(),
                'focalPersonAName': request.form.get('focalPersonAName', '').strip(),
                'focalPersonAPosition': request.form.get('focalPersonAPosition', '').strip(),
                'focalPersonAPhone': request.form.get('focalPersonAPhone', '').strip(),
                'focalPersonAEmail': request.form.get('focalPersonAEmail', '').strip(),
                'agreementStartDate': request.form.get('agreementStartDate', '').strip(),
                'agreementEndDate': request.form.get('agreementEndDate', '').strip(),
                'totalFeeUSD': request.form.get('totalFeeUSD', '').strip(),
                'totalFeeWords': request.form.get('totalFeeWords', '').strip(),
                'paymentInstallmentDesc': request.form.getlist('paymentInstallmentDesc[]'),
                'workshopDescription': request.form.get('workshopDescription', '').strip(),
                'deliverables': request.form.get('deliverables', '').strip(),
                'articles': [{'article_number': num.strip(), 'custom_sentence': sent.strip()} for num, sent in zip(request.form.getlist('articleNumber[]'), request.form.getlist('customSentence[]')) if sent.strip()],
                'partyBSignatureNameConfirm': request.form.get('partyBSignatureNameConfirm', '').strip(),
                'title': request.form.get('title', '').strip()
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
                flash('Contract number already exists.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            # Date validation
            start_date = form_data['agreementStartDate']
            end_date = form_data['agreementEndDate']
            if start_date and end_date:
                try:
                    if datetime.strptime(end_date, '%Y-%m-%d') < datetime.strptime(start_date, '%Y-%m-%d'):
                        flash('End date cannot be before start date.', 'danger')
                        return render_template('contracts/create.html', form_data=form_data)
                except ValueError:
                    flash('Invalid date format.', 'danger')
                    return render_template('contracts/create.html', form_data=form_data)

            # Total fee validation
            try:
                total_fee_usd = float(form_data['totalFeeUSD'])
                if total_fee_usd < 0:
                    flash('Total fee cannot be negative.', 'danger')
                    return render_template('contracts/create.html', form_data=form_data)
            except ValueError:
                flash('Invalid total fee amount.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            # Tax percentage validation
            try:
                tax_percentage = float(form_data['taxPercentage']) if form_data['taxPercentage'] else 15.0
                if tax_percentage < 0:
                    flash('Tax percentage cannot be negative.', 'danger')
                    return render_template('contracts/create.html', form_data=form_data)
            except ValueError:
                flash('Invalid tax percentage.', 'danger')
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
                    article_num = int(article['article_number'])
                    if article['custom_sentence'] and article_num not in article_numbers:
                        articles.append(article)
                        article_numbers.add(article_num)
                    elif article['custom_sentence'] and article_num in article_numbers:
                        flash(f'Duplicate article number {article_num} detected.', 'danger')
                        return render_template('contracts/create.html', form_data=form_data)
                except ValueError:
                    flash(f'Invalid article number: {article["article_number"]}.', 'danger')
                    return render_template('contracts/create.html', form_data=form_data)

            # Deliverables processing
            deliverables = form_data['deliverables'].split('\n')
            deliverables = [d.strip() for d in deliverables if d.strip()]
            if not deliverables:
                flash('At least one deliverable is required.', 'danger')
                return render_template('contracts/create.html', form_data=form_data)

            # Create total_fee_words if not provided
            total_fee_words = form_data['totalFeeWords'] or number_to_words(total_fee_usd)

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
                payment_installment_desc='; '.join(payment_installments) if payment_installments else '',
                payment_gross=f"USD {total_fee_usd:.2f}" if total_fee_usd else '',
                payment_net=f"USD {total_fee_usd * (1 - tax_percentage / 100):.2f}" if total_fee_usd and tax_percentage is not None else '',
                workshop_description=form_data['workshopDescription'],
                focal_person_a_name=form_data['focalPersonAName'],
                focal_person_a_position=form_data['focalPersonAPosition'],
                focal_person_a_phone=form_data['focalPersonAPhone'],
                focal_person_a_email=form_data['focalPersonAEmail'],
                party_a_signature_name=request.form.get('partyASignatureName', 'Mr. SOEUNG Saroeun'),
                party_b_signature_name=form_data['partyBSignatureName'],
                party_b_position=form_data['partyBPosition'],
                total_fee_words=total_fee_words,
                title=form_data['title'],
                deliverables='; '.join(deliverables),
                output_description=form_data['outputDescription'],
                custom_article_sentences={article['article_number']: article['custom_sentence'] for article in articles}
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

# View contract
@contracts_bp.route('/view/<string:contract_id>')
@login_required
def view(contract_id):
    try:
        # Fetch contract
        contract = Contract.query.get_or_404(contract_id)
        contract_data = contract.to_dict()

        # Prepare dynamic data
        total_fee = float(contract_data['total_fee_usd'] or 0.0)
        tax_percentage = float(contract_data['tax_percentage'] or 15.0)
        tax_amount = total_fee * (tax_percentage / 100)
        net_amount = total_fee - tax_amount
        total_fee_words = contract_data['total_fee_words'] or number_to_words(total_fee)
        deliverables = normalize_to_list(contract_data['deliverables'])

        # Process payment installments
        payment_installments = normalize_to_list(contract_data['payment_installment_desc'])
        installment_data = []
        total_percentage = 0
        for desc in payment_installments:
            match = re.search(r'\((\d+)%\)', desc)
            perc = int(match.group(1)) if match else 0
            total_percentage += perc
            amount = total_fee * (perc / 100)
            tax_am = amount * (tax_percentage / 100)
            net_am = amount - tax_am
            installment_data.append({
                'desc': desc,
                'amount': f"{amount:.2f}",
                'tax_am': f"{tax_am:.2f}",
                'net_am': f"{net_am:.2f}",
                'perc': perc
            })

        if total_percentage != 100 and payment_installments:
            logger.warning(f"Total percentages for contract {contract_id} do not add up to 100%: {total_percentage}")

        # Process custom articles
        articles = contract_data['articles'] if contract_data.get('articles') else []

        # Define standard articles to match Mr.Sean Bunrith.docx
        standard_articles = [
            {
                'number': 1,
                'title': 'TERMS OF REFERENCE',
                'content': '"Party B" shall perform tasks as stated in the attached TOR (annex-1) to "Party A", and deliver each milestone as stipulated in article 4.\nThe work shall be of good quality and well performed with the acceptance by "Party A".'
            },
            {
                'number': 2,
                'title': 'TERM OF AGREEMENT',
                'content': f'This agreement is effective from {format_date(contract_data["agreement_start_date"])} – {format_date(contract_data["agreement_end_date"])}.\nThis Agreement is terminated automatically after the due date of the Agreement Term unless otherwise, both Parties agree to extend the Term with a written agreement.'
            },
            {
                'number': 3,
                'title': 'PROFESSIONAL FEE',
                'content': f'The professional fee is the total amount of USD {total_fee:.2f} ({total_fee_words}) including tax for the whole assignment period.\nTotal Service Fee: USD {total_fee:.2f}\nWithholding Tax {tax_percentage}%: USD {tax_amount:.2f}\nNet amount: USD {net_amount:.2f}\n"Party B" is responsible to issue the Invoice (net amount) and receipt (when receiving the payment) with the total amount as stipulated in each instalment as in the Article 4 after having done the agreed deliverable tasks, for payment request.\nThe payment will be processed after the satisfaction from "Party A" as of the required deliverable tasks as stated in Article 4.\n"Party B" is responsible for all related taxes payable to the government department.'
            },
            {
                'number': 4,
                'title': 'TERM OF PAYMENT',
                'content': 'The payment will be made based on the following schedules:',
                'table': [
                    {'Installment': 'Installment', 'Total Amount (USD)': 'Total Amount (USD)', 'Deliverable': 'Deliverable', 'Due date': 'Due date'},
                    *[{
                        'Installment': inst['desc'],
                        'Total Amount (USD)': f'· Gross: ${inst["amount"]}\n· Tax {tax_percentage}%: ${inst["tax_am"]}\n· Net pay: ${inst["net_am"]}',
                        'Deliverable': '\n'.join([f'· {d}' for d in deliverables]),
                        'Due date': format_date(contract_data['agreement_end_date'])
                    } for inst in installment_data]
                ]
            },
            {
                'number': 5,
                'title': 'NO OTHER PERSONS',
                'content': 'No person or entity, which is not a party to this agreement, has any rights to enforce, take any action, or claim it is owed any benefit under this agreement.'
            },
            {
                'number': 6,
                'title': 'MONITORING and COORDINATION',
                'content': f'{contract_data["focal_person_a_name"] or "N/A"}, {contract_data["focal_person_a_position"] or "N/A"} (Telephone {contract_data["focal_person_a_phone"] or "N/A"} Email: {contract_data["focal_person_a_email"] or "N/A"}) is the focal contact person of "Party A" and {contract_data["party_b_signature_name"] or "N/A"} (HP. {contract_data["party_b_phone"] or "N/A"}, E-mail: {contract_data["party_b_email"] or "N/A"}) the focal contact person of "Party B".\nThe focal contact person of "Party A" and "Party B" will work together for overall coordination including reviewing and meeting discussions during the assignment process.'
            },
            {
                'number': 7,
                'title': 'CONFIDENTIALITY',
                'content': f'All outputs produced, with the exception of the “{contract_data["output_description"] or "N/A"}”, which is a contribution from, and to be claimed as a public document by the main author and co-author in associated, and/or under this agreement, shall be the property of "Party A".\nThe "Party B" agrees to not disclose any confidential information, of which he/she may take cognizance in the performance under this contract, except with the prior written approval of the "Party A".'
            },
            {
                'number': 8,
                'title': 'ANTI-CORRUPTION and CONFLICT OF INTEREST',
                'content': '"Party B" shall not participate in any practice that is or could be construed as an illegal or corrupt practice in Cambodia.\nThe "Party A" is committed to fighting all types of corruption and expects this same commitment from the consultant it reserves the rights and believes based on the declaration of "Party B" that it is an independent social enterprise firm operating in Cambodia and it does not involve any conflict of interest with other parties that may be affected to the "Party A".'
            },
            {
                'number': 9,
                'title': 'OBLIGATION TO COMPLY WITH THE NGOF’S POLICIES AND CODE OF CONDUCT',
                'content': 'By signing this agreement, "Party B" is obligated to comply with and respect all existing policies and code of conduct of "Party A", such as Gender Mainstreaming, Child Protection, Disability policy, Environmental Mainstreaming, etc. and the "Party B" declared themselves that s/he will perform the assignment in the neutral position, professional manner, and not be involved in any political affiliation.'
            },
            {
                'number': 10,
                'title': 'ANTI-TERRORISM FINANCING AND FINANCIAL CRIME',
                'content': 'NGOF is determined that all its funds and resources should only be used to further its mission and shall not be subject to illicit use by any third party nor used or abused for any illicit purpose.\nIn order to achieve this objective, NGOF will not knowingly or recklessly provide funds, economic goods, or material support to any entity or individual designated as a “terrorist” by the international community or affiliate domestic governments and will take all reasonable steps to safeguard and protect its assets from such illicit use and to comply with host government laws.\nNGOF respects its contracts with its donors and puts procedures in place for compliance with these contracts.\n“Illicit use” refers to terrorist financing, sanctions, money laundering, and export control regulations.'
            },
            {
                'number': 11,
                'title': 'INSURANCE',
                'content': '"Party B" is responsible for any health and life insurance of its team members. "Party A" will not be held responsible for any medical expenses or compensation incurred during or after this contract.'
            },
            {
                'number': 12,
                'title': 'ASSIGNMENT',
                'content': '"Party B" shall have the right to assign individuals within its organization to carry out the tasks herein named in the attached Technical Proposal.\nThe "Party B" shall not assign, or transfer any of its rights or obligations under this agreement hereunder without the prior written consent of "Party A".\nAny attempt by "Party B" to assign or transfer any of its rights and obligations without the prior written consent of "Party A" shall render this agreement subject to immediate termination by "Party A".'
            },
            {
                'number': 13,
                'title': 'RESOLUTION OF CONFLICTS/DISPUTES',
                'content': 'Conflicts between any of these agreements shall be resolved by the following methods:\nIn the case of a disagreement arising between "Party A" and the "Party B" regarding the implementation of any part of, or any other substantive question arising under or relating to this agreement, the parties shall use their best efforts to arrive at an agreeable resolution by mutual consultation.\nUnresolved issues may, upon the option of either "party a"nd written notice to the other party, be referred to for arbitration.\nFailure by the "Party B" or "Party A" to dispute a decision arising from such arbitration in writing within thirty (30) calendar days of receipt of a final decision shall result in such final decision being deemed binding upon either the "Party B" and/or "Party A".\nAll expenses related to arbitration will be shared equally between both parties.'
            },
            {
                'number': 14,
                'title': 'TERMINATION',
                'content': 'The "Party A" or the "Party B" may, by notice in writing, terminate this agreement under the following conditions:\n1. "Party A" may terminate this agreement at any time with a week notice if "Party B" fails to comply with the terms and conditions of this agreement.\n2. For gross professional misconduct (as defined in the NGOF Human Resource Policy), "Party A" may terminate this agreement immediately without prior notice.\n"Party A" will notify "Party B" in a letter that will indicate the reason for termination as well as the effective date of termination.\n3. "Party B" may terminate this agreement at any time with a one-week notice if "Party A" fails to comply with the terms and conditions of this agreement.\n"Party B" will notify "Party A" in a letter that will indicate the reason for termination as well as the effective date of termination.\nBut if "Party B" intended to terminate this agreement by itself without any appropriate reason or fails of implementing the assignment, "Party B" has to refund the full amount of fees received to "Party A".\n4. If for any reason either "Party A" or the "Party B" decides to terminate this agreement, "Party B" shall be paid pro-rata for the work already completed by "Party A".\nThis payment will require the submission of a timesheet that demonstrates work completed as well as the handing over of any deliverables completed or partially completed.\nIn case "Party B" has received payment for services under the agreement which have not yet been performed; the appropriate portion of these fees would be refunded by "Party B" to "Party A".'
            },
            {
                'number': 15,
                'title': 'MODIFICATION OR AMENDMENT',
                'content': 'No modification or amendment of this agreement shall be valid unless in writing and signed by an authorized person of "Party A" and "Party B".'
            },
            {
                'number': 16,
                'title': 'CONTROLLING OF LAW',
                'content': 'This agreement shall be governed and construed following the law of the Kingdom of Cambodia.\nThe Simultaneous Interpretation Agreement is prepared in two original copies.'
            }
        ]

        return render_template(
            'contracts/view.html',
            contract=contract_data,
            total_fee=total_fee,
            tax_percentage=tax_percentage,
            tax_amount=tax_amount,
            net_amount=net_amount,
            total_fee_words=total_fee_words,
            deliverables=deliverables,
            installment_data=installment_data,
            articles=articles,
            standard_articles=standard_articles,
            format_date=format_date
        )

    except Exception as e:
        logger.error(f"Error viewing contract {contract_id}: {str(e)}")
        flash(f"Error viewing contract: {str(e)}", 'danger')
        return redirect(url_for('contracts.index'))

# Update contract
@contracts_bp.route('/update/<string:contract_id>', methods=['GET', 'POST'])
@login_required
def update(contract_id):
    try:
        contract = Contract.query.get_or_404(contract_id)
        form_data = {}
        if request.method == 'POST':
            try:
                # Collect all form data for re-rendering on error
                form_data = {
                    'id': contract_id,
                    'projectTitle': request.form.get('projectTitle', '').strip(),
                    'contractNumber': request.form.get('contractNumber', '').strip(),
                    'outputDescription': request.form.get('outputDescription', '').strip(),
                    'taxPercentage': request.form.get('taxPercentage', '').strip(),
                    'partyBSignatureName': request.form.get('partyBSignatureName', '').strip(),
                    'partyBPosition': request.form.get('partyBPosition', '').strip(),
                    'partyBPhone': request.form.get('partyBPhone', '').strip(),
                    'partyBEmail': request.form.get('partyBEmail', '').strip(),
                    'partyBAddress': request.form.get('partyBAddress', '').strip(),
                    'focalPersonAName': request.form.get('focalPersonAName', '').strip(),
                    'focalPersonAPosition': request.form.get('focalPersonAPosition', '').strip(),
                    'focalPersonAPhone': request.form.get('focalPersonAPhone', '').strip(),
                    'focalPersonAEmail': request.form.get('focalPersonAEmail', '').strip(),
                    'agreementStartDate': request.form.get('agreementStartDate', '').strip(),
                    'agreementEndDate': request.form.get('agreementEndDate', '').strip(),
                    'totalFeeUSD': request.form.get('totalFeeUSD', '').strip(),
                    'totalFeeWords': request.form.get('totalFeeWords', '').strip(),
                    'paymentInstallmentDesc': request.form.getlist('paymentInstallmentDesc[]'),
                    'workshopDescription': request.form.get('workshopDescription', '').strip(),
                    'deliverables': request.form.get('deliverables', '').strip(),
                    'articles': [{'article_number': num.strip(), 'custom_sentence': sent.strip()} for num, sent in zip(request.form.getlist('articleNumber[]'), request.form.getlist('customSentence[]')) if sent.strip()],
                    'partyBSignatureNameConfirm': request.form.get('partyBSignatureNameConfirm', '').strip(),
                    'title': request.form.get('title', '').strip()
                }

                # Required field validation
                if not form_data['projectTitle']:
                    flash('Project title is required.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)
                if not form_data['contractNumber']:
                    flash('Contract number is required.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)
                if not form_data['partyBSignatureName']:
                    flash('Party B signature name is required.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)
                if not form_data['outputDescription']:
                    flash('Output description is required.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)
                if not form_data['agreementStartDate']:
                    flash('Agreement start date is required.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)
                if not form_data['agreementEndDate']:
                    flash('Agreement end date is required.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)
                if not form_data['totalFeeUSD']:
                    flash('Total fee USD is required.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)
                if not form_data['deliverables']:
                    flash('Deliverables are required.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)
                if not form_data['partyBSignatureNameConfirm']:
                    flash('Party B signature name confirmation is required.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)

                # Signature confirmation validation
                if form_data['partyBSignatureName'] != form_data['partyBSignatureNameConfirm']:
                    flash('Party B Signature Name and Confirmation do not match.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)

                # Unique contract number validation
                existing_contract = Contract.query.filter_by(contract_number=form_data['contractNumber']).first()
                if existing_contract and existing_contract.id != contract_id:
                    flash('Contract number already exists.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)

                # Date validation
                start_date = form_data['agreementStartDate']
                end_date = form_data['agreementEndDate']
                if start_date and end_date:
                    try:
                        if datetime.strptime(end_date, '%Y-%m-%d') < datetime.strptime(start_date, '%Y-%m-%d'):
                            flash('End date cannot be before start date.', 'danger')
                            return render_template('contracts/update.html', form_data=form_data)
                    except ValueError:
                        flash('Invalid date format.', 'danger')
                        return render_template('contracts/update.html', form_data=form_data)

                # Total fee validation
                try:
                    total_fee_usd = float(form_data['totalFeeUSD'])
                    if total_fee_usd < 0:
                        flash('Total fee cannot be negative.', 'danger')
                        return render_template('contracts/update.html', form_data=form_data)
                except ValueError:
                    flash('Invalid total fee amount.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)

                # Tax percentage validation
                try:
                    tax_percentage = float(form_data['taxPercentage']) if form_data['taxPercentage'] else 15.0
                    if tax_percentage < 0:
                        flash('Tax percentage cannot be negative.', 'danger')
                        return render_template('contracts/update.html', form_data=form_data)
                except ValueError:
                    flash('Invalid tax percentage.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)

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
                            return render_template('contracts/update.html', form_data=form_data)
                    except (IndexError, ValueError):
                        flash(f'Invalid format for installment description: {desc}. Please use format like "Installment #1 (50%)".', 'danger')
                        return render_template('contracts/update.html', form_data=form_data)

                # Article validation
                articles = []
                article_numbers = set()
                for article in form_data['articles']:
                    try:
                        article_num = int(article['article_number'])
                        if article['custom_sentence'] and article_num not in article_numbers:
                            articles.append(article)
                            article_numbers.add(article_num)
                        elif article['custom_sentence'] and article_num in article_numbers:
                            flash(f'Duplicate article number {article_num} detected.', 'danger')
                            return render_template('contracts/update.html', form_data=form_data)
                    except ValueError:
                        flash(f'Invalid article number: {article["article_number"]}.', 'danger')
                        return render_template('contracts/update.html', form_data=form_data)

                # Deliverables processing
                deliverables = form_data['deliverables'].split('\n')
                deliverables = [d.strip() for d in deliverables if d.strip()]
                if not deliverables:
                    flash('At least one deliverable is required.', 'danger')
                    return render_template('contracts/update.html', form_data=form_data)

                # Create total_fee_words if not provided
                total_fee_words = form_data['totalFeeWords'] or number_to_words(total_fee_usd)

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
                contract.payment_installment_desc = '; '.join(payment_installments) if payment_installments else ''
                contract.payment_gross = f"USD {total_fee_usd:.2f}" if total_fee_usd else ''
                contract.payment_net = f"USD {total_fee_usd * (1 - tax_percentage / 100):.2f}" if total_fee_usd and tax_percentage is not None else ''
                contract.workshop_description = form_data['workshopDescription']
                contract.focal_person_a_name = form_data['focalPersonAName']
                contract.focal_person_a_position = form_data['focalPersonAPosition']
                contract.focal_person_a_phone = form_data['focalPersonAPhone']
                contract.focal_person_a_email = form_data['focalPersonAEmail']
                contract.party_a_signature_name = request.form.get('partyASignatureName', contract.party_a_signature_name or 'Mr. SOEUNG Saroeun')
                contract.party_b_signature_name = form_data['partyBSignatureName']
                contract.party_b_position = form_data['partyBPosition']
                contract.total_fee_words = total_fee_words
                contract.title = form_data['title']
                contract.deliverables = '; '.join(deliverables)
                contract.output_description = form_data['outputDescription']
                contract.custom_article_sentences = {article['article_number']: article['custom_sentence'] for article in articles}

                db.session.commit()
                flash('Contract updated successfully!', 'success')
                return redirect(url_for('contracts.index'))
            except Exception as e:
                db.session.rollback()
                logger.error(f'Error updating contract {contract_id}: {str(e)}')
                flash(f'Error updating contract: {str(e)}', 'danger')
                return render_template('contracts/update.html', form_data=form_data)

        # Prepare form_data for GET request
        contract_data = contract.to_dict()
        contract_data['agreement_start_date_display'] = format_date(contract_data['agreement_start_date'])
        contract_data['agreement_end_date_display'] = format_date(contract_data['agreement_end_date'])
        contract_data['paymentInstallmentDesc'] = normalize_to_list(contract_data['payment_installment_desc'])
        contract_data['deliverables'] = '\n'.join(normalize_to_list(contract_data['deliverables'])) if contract_data['deliverables'] else ''
        contract_data['totalFeeUSD'] = f"{contract_data['total_fee_usd']:.2f}" if contract_data['total_fee_usd'] is not None else ''
        contract_data['totalFeeWords'] = contract_data['total_fee_words'] or ''
        contract_data['projectTitle'] = contract_data['project_title']
        contract_data['contractNumber'] = contract_data['contract_number']
        contract_data['outputDescription'] = contract_data['output_description']
        contract_data['taxPercentage'] = f"{contract_data['tax_percentage']:.2f}" if contract_data['tax_percentage'] is not None else ''
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
        return render_template('contracts/update.html', form_data=contract_data)

    except Exception as e:
        logger.error(f"Error loading update form for contract {contract_id}: {str(e)}")
        flash(f"Error loading contract: {str(e)}", 'danger')
        return redirect(url_for('contracts.index'))

# Delete contract
@contracts_bp.route('/delete/<string:contract_id>', methods=['POST'])
@login_required
def delete(contract_id):
    try:
        contract = Contract.query.get_or_404(contract_id)
        db.session.delete(contract)
        db.session.commit()
        flash('Contract deleted successfully!', 'success')
    except Exception as e:
        db.session.rollback()
        logger.error(f'Error deleting contract {contract_id}: {str(e)}')
        flash(f'Error deleting contract: {str(e)}', 'danger')
    return redirect(url_for('contracts.index'))

# Export contracts to Excel or CSV
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
            contract['articles'] = '; '.join([f"Article {a['article_number']}: {a['custom_sentence']}" for a in contract['articles']]) if contract['articles'] else 'N/A'

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
            'Deliverables': '; '.join(normalize_to_list(c['deliverables'])) if c['deliverables'] else '',
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


# Export contract to DOCX
@contracts_bp.route('/export_docx/<string:contract_id>')
@login_required
def export_docx(contract_id):
    try:
        # Fetch contract
        contract = Contract.query.get_or_404(contract_id)
        contract_data = contract.to_dict()

        # Prepare dynamic data
        total_fee = float(contract_data['total_fee_usd'] or 0.0)
        tax_percentage = float(contract_data['tax_percentage'] or 15.0)
        tax_amount = total_fee * (tax_percentage / 100)
        net_amount = total_fee - tax_amount
        total_fee_words = contract_data['total_fee_words'] or number_to_words(total_fee)
        deliverables = normalize_to_list(contract_data['deliverables'])

        # Process payment installments
        payment_installments = normalize_to_list(contract_data['payment_installment_desc'])
        installment_data = []
        total_percentage = 0
        for desc in payment_installments:
            match = re.search(r'\((\d+)%\)', desc)
            perc = int(match.group(1)) if match else 0
            total_percentage += perc
            amount = total_fee * (perc / 100)
            tax_am = amount * (tax_percentage / 100)
            net_am = amount - tax_am
            installment_data.append({
                'desc': desc,
                'amount': f"{amount:.2f}",
                'tax_am': f"{tax_am:.2f}",
                'net_am': f"{net_am:.2f}",
                'perc': perc
            })

        if total_percentage != 100 and payment_installments:
            logger.warning(f"Total percentages for contract {contract_id} do not add up to 100%: {total_percentage}")

        # Process custom articles
        articles = contract_data['articles'] if contract_data.get('articles') else []

        # Create a new Document
        doc = Document()

        # Set document styles to match Calibri font as in template
        doc.styles['Normal'].font.name = 'Calibri'
        doc.styles['Normal'].font.size = Pt(11)

        # Helper function to add centered paragraph
        def add_centered_paragraph(text, bold=False, font_size=11):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            run.bold = bold
            run.font.name = 'Calibri'
            run.font.size = Pt(font_size)
            return p

        # Helper function to add regular paragraph
        def add_paragraph(text, bold=False, font_size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT):
            p = doc.add_paragraph()
            p.alignment = alignment
            run = p.add_run(text)
            run.bold = bold
            run.font.name = 'Calibri'
            run.font.size = Pt(font_size)
            return p

        # Helper function to add article heading
        def add_article_heading(number, title):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run("ARTICLE")
            run.bold = True
            run.underline = True
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
            run = p.add_run(f" {number}: {title}")
            run.bold = True
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
            return p

        # Helper function to add paragraph with selective bold and underline
        def add_styled_paragraph(text, bold_terms=None, underline_terms=None, alignment=WD_ALIGN_PARAGRAPH.LEFT, font_size=11):
            p = doc.add_paragraph()
            p.alignment = alignment
            # Split text into parts to handle bold and underline separately
            parts = [text]
            if bold_terms:
                parts = []
                current_text = text
                for term in bold_terms:
                    split_text = current_text.split(term)
                    for i, piece in enumerate(split_text[:-1]):
                        parts.append(piece)
                        parts.append(term)
                    parts.append(split_text[-1])
                    current_text = ''.join(parts[-len(split_text)+1:])
            for part in parts:
                run = p.add_run(part)
                run.font.name = 'Calibri'
                run.font.size = Pt(font_size)
                if bold_terms and part in bold_terms:
                    run.bold = True
                if underline_terms and part in underline_terms:
                    run.underline = True
                    run.font.color.rgb = RGBColor(0, 0, 255)
            return p

        # Add header
        add_centered_paragraph("The Service Agreement", bold=True, font_size=14)
        add_centered_paragraph("ON", bold=True)
        add_centered_paragraph(contract_data['project_title'] or 'N/A', font_size=12)
        add_centered_paragraph(f"No.: {contract_data['contract_number'] or 'N/A'}", bold=True)
        add_centered_paragraph("BETWEEN")

        # Party A
        party_a_text = (f"The NGO Forum on Cambodia, represented by {contract_data['party_a_name'] or 'Mr. Soeung Saroeun'}, "
                        f"{contract_data['party_a_position'] or 'Executive Director'}.\n"
                        f"Address: {contract_data['party_a_address'] or '#9-11, Street 476, Sangkat Tuol Tumpoung I, Phnom Penh, Cambodia'}.\n"
                        "hereinafter called the “Party A”")
        add_styled_paragraph(party_a_text, bold_terms=["Party A"])

        add_centered_paragraph("AND")

        # Party B
        party_b_text = (f"{contract_data['party_b_position'] or 'N/A'} {contract_data['party_b_signature_name'] or 'N/A'},\n"
                        f"Address: {contract_data['party_b_address'] or 'N/A'}\n"
                        f"H/P: {contract_data['party_b_phone'] or 'N/A'}, E-mail: {contract_data['party_b_email'] or 'N/A'}\n"
                        "hereinafter called the “Party B”")
        add_styled_paragraph(party_b_text, bold_terms=["Party B", contract_data['party_b_signature_name'] or 'N/A'],
                            underline_terms=[contract_data['party_b_email'] or 'N/A'])

        # Whereas clauses
        add_paragraph(f"Whereas NGOF is a legal entity registered with the Ministry of Interior (MOI) "
                      f"{contract_data['registration_number'] or '#304 សជណ'} dated "
                      f"{contract_data['registration_date'] or '07 March 2012'}.")
        add_styled_paragraph(f"Whereas NGOF will engage the services of “Party B” which accepts the engagement under the following terms and conditions.",
                            bold_terms=["Party B"])
        add_centered_paragraph("Both Parties Agreed as follows:", bold=True)

        # Define standard articles
        standard_articles = [
            {
                'number': 1,
                'title': 'TERMS OF REFERENCE',
                'content': ('Party B shall perform tasks as stated in the attached TOR (annex-1) to Party A, and deliver each milestone as stipulated in article 4.\n'
                            'The work shall be of good quality and well performed with the acceptance by Party A.')
            },
            {
                'number': 2,
                'title': 'TERM OF AGREEMENT',
                'content': (f'The agreement is effective from {format_date(contract_data["agreement_start_date"])} – '
                            f'{format_date(contract_data["agreement_end_date"])}.\n'
                            'This Agreement is terminated automatically after the due date of the Agreement Term unless otherwise, '
                            'both Parties agree to extend the Term with a written agreement.')
            },
            {
                'number': 3,
                'title': 'PROFESSIONAL FEE',
                'content': (f'The professional fee is the total amount of USD {total_fee:.2f} ({total_fee_words}) including tax for the whole assignment period.\n'
                            f'Total Service Fee: USD {total_fee:.2f}\n'
                            f'Withholding Tax {tax_percentage}%: USD {tax_amount:.2f}\n'
                            f'Net amount: USD {net_amount:.2f}\n'
                            'Party B is responsible to issue the Invoice (net amount) and receipt (when receiving the payment) with the total amount as stipulated in each instalment as in the Article 4 after having done the agreed deliverable tasks, for payment request.\n'
                            'The payment will be processed after the satisfaction from Party A as of the required deliverable tasks as stated in Article 4.\n'
                            'Party B is responsible for all related taxes payable to the government department.')
            },
            {
                'number': 4,
                'title': 'TERM OF PAYMENT',
                'content': 'The payment will be made based on the following schedules:',
                'table': [
                    {'Installment': 'Installment', 'Total Amount (USD)': 'Total Amount (USD)', 'Deliverable': 'Deliverable', 'Due date': 'Due date'},
                    *[{
                        'Installment': inst['desc'],
                        'Total Amount (USD)': f'· Gross: ${inst["amount"]}\n· Tax {tax_percentage}%: ${inst["tax_am"]}\n· Net pay: ${inst["net_am"]}',
                        'Deliverable': '\n'.join([f'· {d}' for d in deliverables]),
                        'Due date': format_date(contract_data['agreement_end_date'])
                    } for inst in installment_data]
                ]
            },
            {
                'number': 5,
                'title': 'NO OTHER PERSONS',
                'content': 'No person or entity, which is not a party to this agreement, has any rights to enforce, take any action, or claim it is owed any benefit under this agreement.'
            },
            {
                'number': 6,
                'title': 'MONITORING and COORDINATION',
                'content': (f'{contract_data["focal_person_a_name"] or "N/A"}, {contract_data["focal_person_a_position"] or "N/A"} '
                            f'(Telephone {contract_data["focal_person_a_phone"] or "N/A"} Email: {contract_data["focal_person_a_email"] or "N/A"}) '
                            f'is the focal contact person of Party A and {contract_data["party_b_signature_name"] or "N/A"} '
                            f'(HP. {contract_data["party_b_phone"] or "N/A"}, E-mail: {contract_data["party_b_email"] or "N/A"}) '
                            f'the focal contact person of Party B.\n'
                            'The focal contact person of Party A and Party B will work together for overall coordination including reviewing and meeting discussions during the assignment process.')
            },
            {
                'number': 7,
                'title': 'CONFIDENTIALITY',
                'content': (f'All outputs produced, with the exception of the “{contract_data["output_description"] or "N/A"}”, '
                            'which is a contribution from, and to be claimed as a public document by the main author and co-author in associated, '
                            'and/or under this agreement, shall be the property of Party A.\n'
                            'The Party B agrees to not disclose any confidential information, of which he/she may take cognizance in the performance under this contract, '
                            'except with the prior written approval of the Party A.')
            },
            {
                'number': 8,
                'title': 'ANTI-CORRUPTION and CONFLICT OF INTEREST',
                'content': ('Party B shall not participate in any practice that is or could be construed as an illegal or corrupt practice in Cambodia.\n'
                            'The Party A is committed to fighting all types of corruption and expects this same commitment from the consultant it reserves the rights '
                            'and believes based on the declaration of Party B that it is an independent social enterprise firm operating in Cambodia '
                            'and it does not involve any conflict of interest with other parties that may be affected to the Party A.')
            },
            {
                'number': 9,
                'title': 'OBLIGATION TO COMPLY WITH THE NGOF’S POLICIES AND CODE OF CONDUCT',
                'content': ('By signing this agreement, Party B is obligated to comply with and respect all existing policies and code of conduct of Party A, '
                            'such as Gender Mainstreaming, Child Protection, Disability policy, Environmental Mainstreaming, etc. '
                            'and the Party B declared themselves that s/he will perform the assignment in the neutral position, professional manner, '
                            'and not be involved in any political affiliation.')
            },
            {
                'number': 10,
                'title': 'ANTI-TERRORISM FINANCING AND FINANCIAL CRIME',
                'content': ('NGOF is determined that all its funds and resources should only be used to further its mission and shall not be subject to illicit use by any third party '
                            'nor used or abused for any illicit purpose.\n'
                            'In order to achieve this objective, NGOF will not knowingly or recklessly provide funds, economic goods, or material support to any entity or individual '
                            'designated as a “terrorist” by the international community or affiliate domestic governments and will take all reasonable steps to safeguard and protect '
                            'its assets from such illicit use and to comply with host government laws.\n'
                            'NGOF respects its contracts with its donors and puts procedures in place for compliance with these contracts.\n'
                            '“Illicit use” refers to terrorist financing, sanctions, money laundering, and export control regulations.')
            },
            {
                'number': 11,
                'title': 'INSURANCE',
                'content': ('Party B is responsible for any health and life insurance of its team members. '
                            'Party A will not be held responsible for any medical expenses or compensation incurred during or after this contract.')
            },
            {
                'number': 12,
                'title': 'ASSIGNMENT',
                'content': ('Party B shall have the right to assign individuals within its organization to carry out the tasks herein named in the attached Technical Proposal.\n'
                            'The Party B shall not assign, or transfer any of its rights or obligations under this agreement hereunder without the prior written consent of Party A.\n'
                            'Any attempt by Party B to assign or transfer any of its rights and obligations without the prior written consent of Party A '
                            'shall render this agreement subject to immediate termination by Party A.')
            },
            {
                'number': 13,
                'title': 'RESOLUTION OF CONFLICTS/DISPUTES',
                'content': ('Conflicts between any of these agreements shall be resolved by the following methods:\n'
                            'In the case of a disagreement arising between Party A and the Party B regarding the implementation of any part of, '
                            'or any other substantive question arising under or relating to this agreement, the parties shall use their best efforts to arrive at an agreeable resolution by mutual consultation.\n'
                            'Unresolved issues may, upon the option of either party and written notice to the other party, be referred to for arbitration.\n'
                            'Failure by the Party B or Party A to dispute a decision arising from such arbitration in writing within thirty (30) calendar days of receipt of a final decision '
                            'shall result in such final decision being deemed binding upon either the Party B and/or Party A.\n'
                            'All expenses related to arbitration will be shared equally between both parties.')
            },
            {
                'number': 14,
                'title': 'TERMINATION',
                'content': ('The Party A or the Party B may, by notice in writing, terminate this agreement under the following conditions:\n'
                            '1. Party A may terminate this agreement at any time with a week notice if Party B fails to comply with the terms and conditions of this agreement.\n'
                            '2. For gross professional misconduct (as defined in the NGOF Human Resource Policy), Party A may terminate this agreement immediately without prior notice.\n'
                            'Party A will notify Party B in a letter that will indicate the reason for termination as well as the effective date of termination.\n'
                            '3. Party B may terminate this agreement at any time with a one-week notice if Party A fails to comply with the terms and conditions of this agreement.\n'
                            'Party B will notify Party A in a letter that will indicate the reason for termination as well as the effective date of termination.\n'
                            'But if Party B intended to terminate this agreement by itself without any appropriate reason or fails of implementing the assignment, '
                            'Party B has to refund the full amount of fees received to Party A.\n'
                            '4. If for any reason either Party A or the Party B decides to terminate this agreement, '
                            'Party B shall be paid pro-rata for the work already completed by Party A.\n'
                            'This payment will require the submission of a timesheet that demonstrates work completed as well as the handing over of any deliverables completed or partially completed.\n'
                            'In case Party B has received payment for services under the agreement which have not yet been performed; '
                            'the appropriate portion of these fees would be refunded by Party B to Party A.')
            },
            {
                'number': 15,
                'title': 'MODIFICATION OR AMENDMENT',
                'content': 'No modification or amendment of this agreement shall be valid unless in writing and signed by an authorized person of Party A and Party B.'
            },
            {
                'number': 16,
                'title': 'CONTROLLING OF LAW',
                'content': ('This agreement shall be governed and construed following the law of the Kingdom of Cambodia.\n'
                            'The Simultaneous Interpretation Agreement is prepared in two original copies.')
            }
        ]

        # Add articles
        for article in standard_articles:
            add_article_heading(article['number'], article['title'])
            bold_terms = ['Party A', 'Party B']
            underline_terms = []
            if article['number'] == 6:
                underline_terms = [contract_data['focal_person_a_email'] or 'N/A', contract_data['party_b_email'] or 'N/A']
            elif article['number'] == 7:
                underline_terms = [f'“{contract_data["output_description"] or "N/A"}”']
            add_styled_paragraph(article['content'], bold_terms=bold_terms, underline_terms=underline_terms)

            # Add table for Article 4
            if article.get('table'):
                table = doc.add_table(rows=len(article['table']), cols=len(article['table'][0]))
                table.style = 'Table Grid'
                for i, row_data in enumerate(article['table']):
                    for j, key in enumerate(row_data.keys()):
                        cell = table.rows[i].cells[j]
                        cell.text = row_data[key]
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                        for run in cell.paragraphs[0].runs:
                            run.font.name = 'Calibri'
                            run.font.size = Pt(11)

            # Add custom sentences for the article
            for custom in articles:
                if custom['article_number'] == str(article['number']):
                    add_paragraph(custom['custom_sentence'])

        # Add signatures
        add_centered_paragraph(f"Date: {format_date(contract_data['agreement_start_date'])}", bold=True)
        table = doc.add_table(rows=2, cols=2)
        table.style = None  # No border
        table.autofit = True
        table.allow_autofit = True
        table.columns[0].width = Inches(3.5)
        table.columns[1].width = Inches(3.5)

        # Party A signature
        cell = table.rows[0].cells[0]
        p = cell.add_paragraph("For “Party A”")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.bold = True
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
        p = cell.add_paragraph("_________________")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = cell.add_paragraph(f"{contract_data['party_a_signature_name'] or 'Mr. SOEUNG Saroeun'}")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.bold = True
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
        p = cell.add_paragraph(f"{contract_data['party_a_position'] or 'Executive Director'}")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.bold = True
            run.font.name = 'Calibri'
            run.font.size = Pt(11)

        # Party B signature
        cell = table.rows[0].cells[1]
        p = cell.add_paragraph("For “Party B”")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.bold = True
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
        p = cell.add_paragraph("_________________")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = cell.add_paragraph(f"{contract_data['party_b_signature_name'] or 'N/A'}")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.bold = True
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
        p = cell.add_paragraph(f"{contract_data['party_b_position'] or 'N/A'}")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.bold = True
            run.font.name = 'Calibri'
            run.font.size = Pt(11)

        # Save document to BytesIO
        output = BytesIO()
        doc.save(output)
        output.seek(0)

        # Send file
        return send_file(
            output,
            download_name=f"{contract_data['party_b_signature_name'] or 'contract'}.docx",
            as_attachment=True
        )

    except Exception as e:
        logger.error(f"Error exporting contract {contract_id} to DOCX: {str(e)}")
        flash(f"Error exporting contract to DOCX: {str(e)}", 'danger')
        return redirect(url_for('contracts.view', contract_id=contract_id))