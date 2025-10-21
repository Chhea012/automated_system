from app import db
from datetime import datetime
import uuid

class Intern(db.Model):
    __tablename__ = 'interns'
    __table_args__ = {'extend_existing': True}

    id = db.Column(db.String(36), primary_key=True, default=lambda: str(uuid.uuid4()))
    intern_name = db.Column(db.String(255), nullable=False)  # Full name including title, e.g., "Ms. Dorn Sochea"
    intern_role = db.Column(db.String(255), nullable=False)  # e.g., Finance and Administrative Intern
    intern_address = db.Column(db.Text, default='')  # Intern's address
    intern_phone = db.Column(db.String(20), default='')  # Intern's phone number
    intern_email = db.Column(db.String(100), default='')  # Intern's email
    start_date = db.Column(db.Date, nullable=False)  # Internship start date
    duration = db.Column(db.String(50), nullable=False)  # e.g., "3 months"
    end_date = db.Column(db.Date, nullable=False)  # Internship end date
    working_hours = db.Column(db.String(100), default='8:00am – 5:00pm, Monday to Friday')
    allowance_amount = db.Column(db.Numeric(10, 2), default=0.0)  # Allowance in USD
    has_nssf = db.Column(db.Boolean, default=False)  # Whether NSSF benefit applies
    supervisor_info = db.Column(db.JSON, default=lambda: {})  # e.g., {"title": "MACOR Program Manager", "name": "Mr. SOM Chettana"}
    employer_representative_name = db.Column(db.String(100), default='Mr. Soeung Saroeun')
    employer_representative_title = db.Column(db.String(100), default='Executive Director')
    employer_address = db.Column(db.Text, default='#9-11, St. 476, Sangkat ToulTompong I, Khan Chamka Morn, Phnom Penh, Cambodia')
    employer_phone = db.Column(db.String(20), default='023 214 429')
    employer_fax = db.Column(db.String(20), default='023 994 063')
    employer_email = db.Column(db.String(100), default='info@ngoforum.org.kh')
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    deleted_at = db.Column(db.DateTime, nullable=True)

    def __repr__(self):
        return f"<Intern {self.intern_name}>"

    @property
    def formatted_created_at(self):
        """Format created_at as '17ᵗʰ September 2025'."""
        if not self.created_at:
            return 'N/A'
        day = self.created_at.day
        month = self.created_at.strftime('%B')
        year = self.created_at.year
        suffix = 'th' if 11 <= day % 100 <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
        superscripts = {"st": "ˢᵗ", "nd": "ⁿᵈ", "rd": "ʳᵈ", "th": "ᵗʰ"}
        return f"{day}{superscripts[suffix]} {month} {year}"

    def to_dict(self):
        """Serialize intern data to a dictionary."""
        supervisor_info = self.supervisor_info if isinstance(self.supervisor_info, dict) else {}
        return {
            'id': self.id or '',
            'intern_name': self.intern_name or '',
            'intern_role': self.intern_role or '',
            'intern_address': self.intern_address or '',
            'intern_phone': self.intern_phone or '',
            'intern_email': self.intern_email or '',
            'start_date': self.start_date.strftime('%Y-%m-%d') if self.start_date else '',
            'duration': self.duration or '',
            'end_date': self.end_date.strftime('%d %B %Y') if self.end_date else '',
            'working_hours': self.working_hours or '',
            'allowance_amount': float(self.allowance_amount) if self.allowance_amount is not None else 0.0,
            'has_nssf': self.has_nssf or False,
            'supervisor_info': supervisor_info,
            'employer_representative_name': self.employer_representative_name or '',
            'employer_representative_title': self.employer_representative_title or '',
            'employer_address': self.employer_address or '',
            'employer_phone': self.employer_phone or '',
            'employer_fax': self.employer_fax or '',
            'employer_email': self.employer_email or '',
            'created_at': self.created_at,
            'formatted_created_at': self.formatted_created_at,
            'deleted_at': self.deleted_at
        }
