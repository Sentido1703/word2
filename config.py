import os


class Config:
    SQLALCHEMY_DATABASE_URI = 'postgresql://postgres:12345678@localhost:5432/sentido'
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    TEMPLATE_DIR = os.getenv('TEMPLATE_DIR', './templates')
    OUTPUT_DIR = os.getenv('OUTPUT_DIR', './output')
