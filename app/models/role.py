from datetime import datetime
from .. import db

# Define the many-to-many association table for roles and permissions
role_permissions = db.Table('role_permissions',
    db.Column('role_id', db.Integer, db.ForeignKey('role.id'), primary_key=True),
    db.Column('permission_id', db.Integer, db.ForeignKey('permission.id'), primary_key=True),
    extend_existing=True
)

class Role(db.Model):
    __tablename__ = "role"
    __table_args__ = {'extend_existing': True}  # Allow table redefinition

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(64), unique=True, nullable=False)
    description = db.Column(db.String(255), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    # Many-to-many relationship with Permission
    permissions = db.relationship('Permission', secondary=role_permissions, backref=db.backref('roles', lazy='dynamic'))

    def __repr__(self):
        return f"<Role {self.name}>"