from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
from .. import db

class User(UserMixin, db.Model):
    __tablename__ = "user"
    __table_args__ = {'extend_existing': True}

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    image = db.Column(db.String(255), nullable=True, default="default_profile.png")
    phone_number = db.Column(db.String(20), unique=True, nullable=True)  # Added unique=True
    address = db.Column(db.String(255), nullable=True)
    role_id = db.Column(db.Integer, db.ForeignKey("role.id"), nullable=True, index=True)
    department_id = db.Column(db.Integer, db.ForeignKey("department.id"), nullable=True, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    role = db.relationship("Role", backref=db.backref("users", lazy="dynamic"), foreign_keys=[role_id], lazy="joined")
    department = db.relationship("Department", backref=db.backref("users", lazy="dynamic"), foreign_keys=[department_id], lazy="joined")

    def set_password(self, password):
        self.password_hash = generate_password_hash(password, method='pbkdf2:sha256:600000')
    
    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def get_image_url(self):
        if self.image and self.image != "default_profile.png":
            return f"/static/uploads/{self.image}"
        return "/static/uploads/default_profile.png"

    def has_role(self, role_name):
        """Check if the user has the specified role."""
        return self.role and self.role.name.lower() == role_name.lower()

    def __repr__(self):
        return f"<User {self.username}>"