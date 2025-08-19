import os
from dotenv import load_dotenv
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from flask_migrate import Migrate
from .config import Config

load_dotenv()

db = SQLAlchemy()
login_manager = LoginManager()
migrate = Migrate()

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    db.init_app(app)
    login_manager.init_app(app)
    migrate.init_app(app, db)

    login_manager.login_view = "auth.login"
    login_manager.login_message_category = "info"

    from .routes.auth import auth_bp
    from .routes.main import main_bp
    from .routes.users import users_bp
    from .routes.permission import permissions_bp

    app.register_blueprint(auth_bp, url_prefix="/auth")
    app.register_blueprint(main_bp)
    app.register_blueprint(users_bp, url_prefix="/users")
    app.register_blueprint(permissions_bp, url_prefix="/permissions")

    from .models.user import User
    from .models.permission import Permission

    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    return app