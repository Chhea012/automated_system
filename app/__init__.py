import os
from dotenv import load_dotenv
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from flask_migrate import Migrate
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from .config import Config

load_dotenv()

db = SQLAlchemy()
login_manager = LoginManager()
migrate = Migrate()
limiter = Limiter(key_func=get_remote_address)

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    db.init_app(app)
    login_manager.init_app(app)
    migrate.init_app(app, db)
    limiter.init_app(app)

    login_manager.login_view = "auth.login"
    login_manager.login_message_category = "info"

    from .routes.auth import auth_bp
    from .routes.main import main_bp
    from .routes.users import users_bp
    from .routes.permission import permissions_bp
    from .routes.role import roles_bp
    from .routes.department import departments_bp
    from .routes.contract import contracts_bp

    app.register_blueprint(auth_bp, url_prefix="/auth")
    app.register_blueprint(main_bp)
    app.register_blueprint(users_bp, url_prefix="/users")
    app.register_blueprint(permissions_bp, url_prefix="/permissions")
    app.register_blueprint(roles_bp, url_prefix="/roles")
    app.register_blueprint(departments_bp, url_prefix="/departments")
    app.register_blueprint(contracts_bp, url_prefix="/contracts")

    from .models.user import User
    from .models.permission import Permission
    from .models.role import Role
    from .models.department import Department
    from .models.contract import Contract
    from .forms import LoginForm, RegisterForm

    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    return app