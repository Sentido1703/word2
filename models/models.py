from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Group(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    course = db.Column(db.String(120), nullable=False)
    students = db.relationship('Student', backref='group', lazy=True)


class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    group_id = db.Column(db.Integer, db.ForeignKey('group.id'), nullable=True)

