from app import create_app
from models.models import db, Group, Student

app = create_app()


def create_tables():
    with app.app_context():
        db.create_all()
        if not Group.query.first():
            group = Group(course="Философы антропологи")
            student1 = Student(name="Илья Репин", email="ilia@primer.com", group=group)
            student2 = Student(name="Михаил Ломоносов", email="misha@primer.com", group=group)
            db.session.add_all([group, student1, student2])
            db.session.commit()


if __name__ == "__main__":
    create_tables()
    app.run(debug=True)
