from app import db
from datetime import datetime
import uuid

class Employee(db.Model):
    __tablename__ = 'employees'
    __table_args__ = {'extend_existing': True}

    # === PRIMARY IDENTIFIER ===
    id = db.Column(db.String(36), primary_key=True, default=lambda: str(uuid.uuid4()))

    # === CONTRACT INFORMATION ===
    contract_no = db.Column(db.String(50), nullable=False, unique=True)
    contract_type = db.Column(db.String(100), nullable=False, default='Fixed Duration Contract (FDC)')

    # === EMPLOYER / ORGANIZATION INFORMATION ===
    organization_name = db.Column(db.String(255), default='The NGO Forum on Cambodia')
    representative_name = db.Column(db.String(100), default='Mr. Soeung Saroeun')
    representative_title = db.Column(db.String(100), default='Executive Director')
    organization_address = db.Column(db.Text, default='#9-11, St. 476, Sangkat Toul Tompong I, Khan Chamka Morn, Phnom Penh, Cambodia')
    organization_tel = db.Column(db.String(50), default='023 214 429')
    organization_fax = db.Column(db.String(50), default='023 994 063')
    organization_email = db.Column(db.String(150), default='info@ngoforum.org.kh')

    # === EMPLOYEE INFORMATION ===
    employee_name = db.Column(db.String(255), nullable=False)
    employee_address = db.Column(db.Text, default='')
    employee_tel = db.Column(db.String(50), default='')
    employee_email = db.Column(db.String(150), default='')
    position_title = db.Column(db.String(150), nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)

    # === WORKING DETAILS ===
    working_hours = db.Column(
        db.String(255),
        default='Monday to Friday: 08:00am-12:00pm and 01:30pm-05:30pm'
    )

    # === SALARY AND PAY INFORMATION ===
    salary_amount = db.Column(db.Numeric(10, 2), default=0.0)
    salary_grade = db.Column(db.String(50), default='')
    salary_amount_words = db.Column(db.String(255), default='')

    # === BENEFITS ===
    medical_allowance = db.Column(db.Numeric(10, 2), default=150.00)
    child_education_allowance = db.Column(db.Numeric(10, 2), default=60.00)
    delivery_benefit = db.Column(db.Numeric(10, 2), default=200.00)
    delivery_benefit_miscarriage = db.Column(db.Numeric(10, 2), default=200.00)
    death_benefit = db.Column(db.Numeric(10, 2), default=200.00)
    severance_percentage = db.Column(db.Numeric(5, 2), default=8.33)
    thirteenth_month_salary = db.Column(db.Boolean, default=True)

    # === SIGNATURES ===
    employer_signature_name = db.Column(db.String(150), default='Mr. Soeung Saroeun')
    employee_signature_name = db.Column(db.String(150), default='')
    employer_signature_date = db.Column(db.Date, nullable=True)
    employee_signature_date = db.Column(db.Date, nullable=True)

    # === SYSTEM META ===
    status = db.Column(db.String(50), default='active')
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)
    deleted_at = db.Column(db.DateTime, nullable=True)

    # === METHODS ===
    def generate_salary_in_words(self):
        """Generate salary amount in words with decimal support."""
        if not self.salary_amount or self.salary_amount <= 0:
            self.salary_amount_words = ''
            return
        
        amount = float(self.salary_amount)
        integer_part = int(amount)
        decimal_part = int(round((amount - integer_part) * 100))
        
        words = self._number_to_words(integer_part) + ' US Dollars'
        if decimal_part > 0:
            words += ' and ' + self._number_to_words(decimal_part) + ' Cents'
        words += ' only'
        self.salary_amount_words = words

    def _number_to_words(self, num):
        """Convert number to words."""
        if num == 0:
            return 'Zero'
        
        ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine']
        teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 
                 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen']
        tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 
                'Seventy', 'Eighty', 'Ninety']
        scales = ['', 'Thousand', 'Million', 'Billion']

        if num < 10:
            return ones[num]
        elif num < 20:
            return teens[num - 10]
        elif num < 100:
            ten = num // 10
            one = num % 10
            return tens[ten] + ('-' + ones[one].lower() if one else '')
        elif num < 1000:
            hundred = num // 100
            rest = num % 100
            return ones[hundred] + ' Hundred' + (' ' + self._number_to_words(rest) if rest else '')
        else:
            parts = []
            scale_index = 0
            while num > 0:
                chunk = num % 1000
                if chunk > 0:
                    parts.insert(0, self._number_to_words(chunk) + (' ' + scales[scale_index] if scales[scale_index] else ''))
                num //= 1000
                scale_index += 1
            return ' '.join(parts).replace('  ', ' ').strip()

    def __repr__(self):
        return f"<Employee {self.employee_name} ({self.contract_no})>"

    @property
    def formatted_created_at(self):
        """Return created_at as '17ᵗʰ September 2025'."""
        if not self.created_at:
            return 'N/A'
        day = self.created_at.day
        month = self.created_at.strftime('%B')
        year = self.created_at.year
        suffix = 'th' if 11 <= day % 100 <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
        superscripts = {"st": "ˢᵗ", "nd": "ⁿᵈ", "rd": "ʳᵈ", "th": "ᵗʰ"}
        return f"{day}{superscripts[suffix]} {month} {year}"

    def to_dict(self):
        """Serialize employee contract data to dictionary."""
        return {
            'id': self.id,
            'contract_no': self.contract_no,
            'contract_type': self.contract_type,
            'organization_name': self.organization_name,
            'representative_name': self.representative_name,
            'representative_title': self.representative_title,
            'organization_address': self.organization_address,
            'organization_tel': self.organization_tel,
            'organization_fax': self.organization_fax,
            'organization_email': self.organization_email,
            'employee_name': self.employee_name,
            'employee_address': self.employee_address,
            'employee_tel': self.employee_tel,
            'employee_email': self.employee_email,
            'position_title': self.position_title,
            'start_date': self.start_date.strftime('%Y-%m-%d') if self.start_date else '',
            'end_date': self.end_date.strftime('%Y-%m-%d') if self.end_date else '',
            'working_hours': self.working_hours,
            'salary_amount': float(self.salary_amount) if self.salary_amount else 0.0,
            'salary_grade': self.salary_grade,
            'salary_amount_words': self.salary_amount_words,
            'medical_allowance': float(self.medical_allowance),
            'child_education_allowance': float(self.child_education_allowance),
            'delivery_benefit': float(self.delivery_benefit),
            'delivery_benefit_miscarriage': float(self.delivery_benefit_miscarriage),
            'death_benefit': float(self.death_benefit),
            'severance_percentage': float(self.severance_percentage),
            'thirteenth_month_salary': self.thirteenth_month_salary,
            'employer_signature_name': self.employer_signature_name,
            'employee_signature_name': self.employee_signature_name,
            'employer_signature_date': self.employer_signature_date.strftime('%Y-%m-%d') if self.employer_signature_date else '',
            'employee_signature_date': self.employee_signature_date.strftime('%Y-%m-%d') if self.employee_signature_date else '',
            'status': self.status,
            'created_at': self.created_at.isoformat() if self.created_at else '',
            'updated_at': self.updated_at.isoformat() if self.updated_at else '',
            'deleted_at': self.deleted_at.isoformat() if self.deleted_at else '',
            'formatted_created_at': self.formatted_created_at
        }