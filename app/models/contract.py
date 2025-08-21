from app import db
import json

class Contract(db.Model):
    __tablename__ = 'contracts'
    __table_args__ = {'extend_existing': True}

    id = db.Column(db.String(36), primary_key=True)
    project_title = db.Column(db.String(255), nullable=False)
    contract_number = db.Column(db.String(50), nullable=False)
    organization_name = db.Column(db.String(100))
    party_a_name = db.Column(db.String(100))
    party_a_position = db.Column(db.String(100))
    party_a_address = db.Column(db.Text)
    party_b_full_name_with_title = db.Column(db.String(255))
    party_b_address = db.Column(db.Text)
    party_b_phone = db.Column(db.String(20))
    party_b_email = db.Column(db.String(100))
    registration_number = db.Column(db.String(50))
    registration_date = db.Column(db.String(50))
    agreement_start_date = db.Column(db.String(50))
    agreement_end_date = db.Column(db.String(50))
    total_fee_usd = db.Column(db.Numeric(10, 2))
    gross_amount_usd = db.Column(db.Numeric(10, 2))
    tax_percentage = db.Column(db.Numeric(5, 2))
    payment_installment_desc = db.Column(db.String(255))
    payment_gross = db.Column(db.String(50))
    payment_net = db.Column(db.String(50))
    workshop_description = db.Column(db.String(255))
    focal_person_a_name = db.Column(db.String(100))
    focal_person_a_position = db.Column(db.String(100))
    focal_person_a_phone = db.Column(db.String(20))
    focal_person_a_email = db.Column(db.String(100))
    party_a_signature_name = db.Column(db.String(100))
    party_b_signature_name = db.Column(db.String(100))
    party_b_position = db.Column(db.String(100))
    total_fee_words = db.Column(db.Text)
    title = db.Column(db.String(255))
    deliverables = db.Column(db.Text)
    output_description = db.Column(db.Text)
    custom_article_sentences = db.Column(db.JSON)

    def __repr__(self):
        return f"<Contract {self.contract_number}>"

    def to_dict(self):
        return {
            'id': self.id,
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
            'registration_number': self.registration_number or '',
            'registration_date': self.registration_date or '',
            'agreement_start_date': self.agreement_start_date or '',
            'agreement_end_date': self.agreement_end_date or '',
            'total_fee_usd': float(self.total_fee_usd) if self.total_fee_usd is not None else 0.0,
            'gross_amount_usd': float(self.gross_amount_usd) if self.gross_amount_usd is not None else 0.0,
            'tax_percentage': float(self.tax_percentage) if self.tax_percentage is not None else 0.0,
            'payment_installment_desc': self.payment_installment_desc or '',
            'payment_gross': self.payment_gross or '',
            'payment_net': self.payment_net or '',
            'workshop_description': self.workshop_description or '',
            'focal_person_a_name': self.focal_person_a_name or '',
            'focal_person_a_position': self.focal_person_a_position or '',
            'focal_person_a_phone': self.focal_person_a_phone or '',
            'focal_person_a_email': self.focal_person_a_email or '',
            'party_a_signature_name': self.party_a_signature_name or '',
            'party_b_signature_name': self.party_b_signature_name or '',
            'party_b_position': self.party_b_position or '',
            'total_fee_words': self.total_fee_words or '',
            'title': self.title or '',
            'deliverables': self.deliverables.split('; ') if self.deliverables else [],
            'output_description': self.output_description or '',
            'custom_article_sentences': json.loads(self.custom_article_sentences) if self.custom_article_sentences and self.custom_article_sentences != '{}' else {}
        }