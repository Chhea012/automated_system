
from app import db
from datetime import datetime
import uuid

class Contract(db.Model):
    __tablename__ = 'contracts'
    __table_args__ = {'extend_existing': True}

    id = db.Column(db.String(36), primary_key=True, default=lambda: str(uuid.uuid4()))
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    project_title = db.Column(db.String(255), nullable=False, default='')
    contract_number = db.Column(db.String(50), nullable=False, default='')
    organization_name = db.Column(db.String(100), default='')
    party_a_name = db.Column(db.String(100), default='')
    party_a_position = db.Column(db.String(100), default='')
    party_a_address = db.Column(db.Text, default='')
    party_b_full_name_with_title = db.Column(db.String(255), default='')
    party_b_address = db.Column(db.Text, default='')
    party_b_phone = db.Column(db.String(20), default='')
    party_b_email = db.Column(db.String(100), default='')
    registration_number = db.Column(db.String(50), default='#304 សជណ')
    registration_date = db.Column(db.String(50), default='07 March 2012')
    agreement_start_date = db.Column(db.String(50), default='')
    agreement_end_date = db.Column(db.String(50), default='')
    total_fee_usd = db.Column(db.Numeric(10, 2), default=0.0)
    gross_amount_usd = db.Column(db.Numeric(10, 2), default=0.0)
    tax_percentage = db.Column(db.Numeric(5, 2), default=15.0)
    payment_gross = db.Column(db.String(50), default='')
    payment_net = db.Column(db.String(50), default='')
    workshop_description = db.Column(db.String(255), default='')
    focal_person_info = db.Column(db.JSON, default=lambda: [])  # New JSON column
    party_a_signature_name = db.Column(db.String(100), default='Mr. SOEUNG Saroeun')
    party_b_signature_name = db.Column(db.String(100), default='')
    party_b_position = db.Column(db.String(100), default='')
    total_fee_words = db.Column(db.Text, default='')
    title = db.Column(db.String(255), default='')
    deliverables = db.Column(db.Text, default='')
    output_description = db.Column(db.Text, default='')
    custom_article_sentences = db.Column(db.JSON, default=lambda: {})
    payment_installments = db.Column(db.JSON, default=lambda: [])
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    deleted_at = db.Column(db.DateTime, nullable=True)

    # Add relationship to User model
    user = db.relationship('User', backref=db.backref('contracts', lazy='dynamic'), lazy='joined')

    def __repr__(self):
        return f"<Contract {self.contract_number} by User {self.user_id}>"

    def to_dict(self):
        custom_sentences = self.custom_article_sentences if isinstance(self.custom_article_sentences, dict) else {}
        payment_installments = self.payment_installments if isinstance(self.payment_installments, list) else []
        focal_person_info = self.focal_person_info if isinstance(self.focal_person_info, list) else []
        return {
            'id': self.id or '',
            'user_id': self.user_id or 0,
            'username': self.user.username if self.user else 'N/A',
            'project_title': self.project_title or '',
            'contract_number': self.contract_number or '',
            'organization_name': self.organization_name or '',
            'party_a_name': self.party_a_name or '',
            'party_a_position': self.party_a_position or '',
            'party_a_address': self.party_a_address or '',
            'party_b_full_name_with_title': self.party_b_full_name_with_title or '',
            'party_b_address': self.party_b_address or '',
            'party_b_phone': self.party_b_phone or '',
            'party_b_email': self.party_b_email or '',
            'registration_number': self.registration_number or '#304 សជណ',
            'registration_date': self.registration_date or '07 March 2012',
            'agreement_start_date': self.agreement_start_date or '',
            'agreement_end_date': self.agreement_end_date or '',
            'total_fee_usd': float(self.total_fee_usd) if self.total_fee_usd is not None else 0.0,
            'gross_amount_usd': float(self.gross_amount_usd) if self.gross_amount_usd is not None else 0.0,
            'tax_percentage': float(self.tax_percentage) if self.tax_percentage is not None else 15.0,
            'payment_installments': payment_installments,
            'payment_gross': self.payment_gross or '',
            'payment_net': self.payment_net or '',
            'workshop_description': self.workshop_description or '',
            'focal_person_info': focal_person_info,
            'party_a_signature_name': self.party_a_signature_name or 'Mr. SOEUNG Saroeun',
            'party_b_signature_name': self.party_b_signature_name or '',
            'party_b_position': self.party_b_position or '',
            'total_fee_words': self.total_fee_words or '',
            'title': self.title or '',
            'deliverables': self.deliverables.split('; ') if self.deliverables else [],
            'output_description': self.output_description or '',
            'custom_article_sentences': custom_sentences,
            'articles': [{'article_number': str(k), 'custom_sentence': v} for k, v in custom_sentences.items()] if custom_sentences else [],
            'created_at': self.created_at,
            'deleted_at': self.deleted_at
        }
