from flask import Flask
from config import Config
from models.models import db
from routes.report_routes import report_bp


def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    db.init_app(app)

    app.register_blueprint(report_bp)

    return app
