from flask import Flask
from flask_cors import CORS
from flask_sqlalchemy import SQLAlchemy
from .models import db
from .security.security import init_jwt
from .routers.auth import auth_bp
from .routers.translation import translation_bp
from .routers.payment import payment_bp
from .routers.history import history_bp
from .routers.admin import admin_bp
from .routers.contact import contact_bp

def create_app(config_class='app.config.DevelopmentConfig'):
    app = Flask(__name__)

    # Load config
    if isinstance(config_class, str):
        app.config.from_object(config_class)
    else:
        app.config.update(config_class)

    # Initialize extensions
    CORS(app)
    db.init_app(app)
    init_jwt(app)

    # Create database tables + run safe migrations
    with app.app_context():
        db.create_all()
        _run_migrations(db)

    # Register blueprints
    app.register_blueprint(auth_bp,        url_prefix='/api/auth')
    app.register_blueprint(translation_bp, url_prefix='/api/translation')
    app.register_blueprint(payment_bp,     url_prefix='/api/payment')
    app.register_blueprint(history_bp,     url_prefix='/api/history')
    app.register_blueprint(admin_bp,       url_prefix='/api/admin')
    app.register_blueprint(contact_bp,     url_prefix='/api/contact')

    return app


def _run_migrations(db):
    """Safe ALTER TABLE migrations for columns/tables added after initial deploy."""
    from sqlalchemy import text

    _MIGRATIONS = [
        # user table
        ("user",        "role",          "ALTER TABLE `user` ADD COLUMN `role` VARCHAR(20) NOT NULL DEFAULT 'user'"),
        ("user",        "token_balance", "ALTER TABLE `user` ADD COLUMN `token_balance` INT NOT NULL DEFAULT 0"),
        # translation table
        ("translation", "type",          "ALTER TABLE `translation` ADD COLUMN `type` VARCHAR(20) NOT NULL DEFAULT 'text'"),
        ("translation", "token_cost",    "ALTER TABLE `translation` ADD COLUMN `token_cost` INT NOT NULL DEFAULT 0"),
        # payment table
        ("payment",     "plan_type",     "ALTER TABLE `payment` ADD COLUMN `plan_type` VARCHAR(50) DEFAULT NULL"),
    ]

    try:
        with db.engine.connect() as conn:
            for table, column, ddl in _MIGRATIONS:
                result = conn.execute(text(
                    "SELECT COUNT(*) FROM information_schema.COLUMNS "
                    "WHERE TABLE_SCHEMA = DATABASE() "
                    "AND TABLE_NAME = :t AND COLUMN_NAME = :c"
                ), {"t": table, "c": column})
                if result.scalar() == 0:
                    conn.execute(text(ddl))
                    conn.commit()
                    print(f"[migration] Added column '{column}' to '{table}'")
    except Exception as e:
        print(f"[migration] Note: {e}")
