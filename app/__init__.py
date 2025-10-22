import os
from dotenv import load_dotenv
from flask import Flask, render_template
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, current_user
from flask_migrate import Migrate
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_mail import Mail
from .config import Config
load_dotenv()
db = SQLAlchemy()
login_manager = LoginManager()
migrate = Migrate()
limiter = Limiter(key_func=get_remote_address)
mail = Mail()
def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)
    # Initialize extensions
    db.init_app(app)
    login_manager.init_app(app)
    migrate.init_app(app, db)
    limiter.init_app(app)
    mail.init_app(app)
    login_manager.login_view = "auth.login"
    login_manager.login_message_category = "info"
    # Register blueprints
    from .routes.auth import auth_bp
    from .routes.main import main_bp
    from .routes.users import users_bp
    from .routes.permission import permissions_bp
    from .routes.role import roles_bp
    from .routes.department import departments_bp
    from .routes.contract import contracts_bp
    from .routes.mydepartments import mydepartments_bp
    from .routes.reports import reports_bp
    from .routes.interns import interns_bp  # Added interns blueprint
    from .routes.employees import employees_bp  # Added employees blueprint
    app.register_blueprint(auth_bp, url_prefix="/auth")
    app.register_blueprint(main_bp)
    app.register_blueprint(users_bp, url_prefix="/users")
    app.register_blueprint(permissions_bp, url_prefix="/permissions")
    app.register_blueprint(roles_bp, url_prefix="/roles")
    app.register_blueprint(departments_bp, url_prefix="/departments")
    app.register_blueprint(contracts_bp, url_prefix="/contracts")
    app.register_blueprint(mydepartments_bp, url_prefix="/mydepartments")
    app.register_blueprint(reports_bp, url_prefix="/reports")
    app.register_blueprint(interns_bp, url_prefix="/interns")  # Register interns blueprint
    app.register_blueprint(employees_bp, url_prefix="/employees")  # Register employees blueprint
    # Import models
    from .models.user import User
    from .models.permission import Permission
    from .models.role import Role
    from .models.department import Department
    from .models.contract import Contract
    from .models.notification import Notification
    from .models.interns import Intern  # Added Intern model
    from .models.employees import Employee  # Added Employee model
    # User loader for Flask-Login
    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))
    # Context processor for notifications
    @app.context_processor
    def inject_notifications():
        if current_user.is_authenticated and current_user.has_role('Admin'):
            unread_count = Notification.query.filter_by(recipient_id=current_user.id, is_read=False).count()
            notifications = Notification.query.filter_by(recipient_id=current_user.id)\
                .order_by(Notification.created_at.desc())\
                .limit(7).all()
            notifications_dict = [notif.to_dict() for notif in notifications]
            return {
                'unread_count': unread_count,
                'notifications': notifications_dict
            }
        return {'unread_count': 0, 'notifications': []}
    # Error handlers
    @app.errorhandler(404)
    def not_found_error(error):
        return render_template('errors/404.html'), 404
    @app.errorhandler(500)
    def internal_error(error):
        db.session.rollback()  # Roll back session to avoid database inconsistencies
        return render_template('errors/500.html'), 500
    return app