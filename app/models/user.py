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
    password_hash = db.Column(db.String(128), nullable=False)

    # profile
    image = db.Column(db.String(255), nullable=True, default="default_profile.png")
    phone_number = db.Column(db.String(20), nullable=True)
    address = db.Column(db.String(255), nullable=True)

    # NEW: relations
    role_id = db.Column(db.Integer, db.ForeignKey("role.id"), nullable=True, index=True)
    department_id = db.Column(db.Integer, db.ForeignKey("department.id"), nullable=True, index=True)

    # timestamps
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # ORM relationships
    role = db.relationship(
        "Role",
        backref=db.backref("users", lazy="dynamic"),
        foreign_keys=[role_id],
        lazy="joined",
    )
    department = db.relationship(
        "Department",
        backref=db.backref("users", lazy="dynamic"),
        foreign_keys=[department_id],
        lazy="joined",
    )

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def get_image_url(self):
        if self.image and self.image != "default_profile.png":
            return f"/static/uploads/{self.image}"
        return "/static/uploads/default_profile.png"

    def __repr__(self):
        return f"<User {self.username}>"
