from datetime import datetime
from .. import db
from .user import User
from .contract import Contract

class Notification(db.Model):
    __tablename__ = 'notifications'
    __table_args__ = {'extend_existing': True}

    id = db.Column(db.Integer, primary_key=True)
    creator_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False, index=True)
    recipient_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False, index=True)
    title = db.Column(db.String(255), nullable=False)
    message = db.Column(db.Text, nullable=False)
    is_read = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    related_contract_id = db.Column(db.String(36), db.ForeignKey('contracts.id'), nullable=True)

    creator = db.relationship('User', foreign_keys=[creator_id], backref=db.backref('sent_notifications', lazy='dynamic'))
    recipient = db.relationship('User', foreign_keys=[recipient_id], backref=db.backref('received_notifications', lazy='dynamic'))
    contract = db.relationship('Contract', backref=db.backref('notifications', lazy='dynamic'))

    def to_dict(self):
        return {
            'id': self.id,
            'title': self.title,
            'message': self.message,
            'is_read': self.is_read,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'creator_username': self.creator.username if self.creator else 'Unknown',
            'creator_image': self.creator.get_image_url() if self.creator else '/static/uploads/default_profile.png',
            'project_title': self.contract.project_title if self.contract else 'N/A',
            'contract_number': self.contract.contract_number if self.contract else 'N/A'
        }

    def __repr__(self):
        return f"<Notification {self.id}: {self.title} for {self.recipient.username}>"