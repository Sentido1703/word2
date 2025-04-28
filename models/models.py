from tortoise import fields, models
from fastadmin import (
    TortoiseModelAdmin,
    WidgetType,
    register,
    action
)
import bcrypt
from enum import Enum


class AdminUser(models.Model):
    """Модель административного пользователя"""
    id = fields.IntField(pk=True)
    username = fields.CharField(max_length=50, unique=True)
    password = fields.CharField(max_length=200)
    is_active = fields.BooleanField(default=True)
    is_superuser = fields.BooleanField(default=False)
    created_at = fields.DatetimeField(auto_now_add=True)
    updated_at = fields.DatetimeField(auto_now=True)

    class Meta:
        table = "admin_users"

    def set_password(self, password: str | bytes):
        if isinstance(password, str):
            self.password = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        else:
            self.password = bcrypt.hashpw(password, bcrypt.gensalt()).decode()


class DayOfWeek(str, Enum):
    MONDAY = "Понедельник"
    TUESDAY = "Вторник"
    WEDNESDAY = "Среда"
    THURSDAY = "Четверг"
    FRIDAY = "Пятница"
    SATURDAY = "Суббота"
    SUNDAY = "Воскресенье"


class Teacher(models.Model):
    """Модель преподавателя"""
    id = fields.IntField(pk=True)
    full_name = fields.CharField(max_length=255, null=False)
    email = fields.CharField(max_length=255, unique=True, null=False)
    # Связи с другими моделями
    groups: fields.ReverseRelation["Group"]
    schedule_items: fields.ReverseRelation["ScheduleItem"]

    class Meta:
        table = "teachers"

    def __str__(self):
        return self.full_name


class Group(models.Model):
    """Модель группы"""
    id = fields.IntField(pk=True)
    code = fields.CharField(max_length=50, null=False)  # Например, КМБО-05-21
    course = fields.CharField(max_length=255, null=False)  # Название курса
    teacher = fields.ForeignKeyField("models.Teacher", related_name="groups", null=True)
    # Связи с другими моделями
    students: fields.ReverseRelation["Student"]
    schedule_items: fields.ReverseRelation["ScheduleItem"]

    class Meta:
        table = "groups"

    def __str__(self):
        return self.code


class Student(models.Model):
    """Модель студента"""
    id = fields.IntField(pk=True)
    full_name = fields.CharField(max_length=255, null=False)
    email = fields.CharField(max_length=255, unique=True, null=False)
    group = fields.ForeignKeyField("models.Group", related_name="students", null=True)
    # Связи с другими моделями
    publications: fields.ReverseRelation["Publication"]
    grades: fields.ReverseRelation["Grade"]

    class Meta:
        table = "students"

    def __str__(self):
        return self.full_name


class Discipline(models.Model):
    """Модель дисциплины/предмета"""
    id = fields.IntField(pk=True)
    name = fields.CharField(max_length=255, null=False)
    education_level = fields.CharField(max_length=100, null=True)  # Бакалавриат/Магистратура
    department = fields.CharField(max_length=255, null=True)
    hours_lecture = fields.IntField(null=True)
    hours_practice = fields.IntField(null=True)
    hours_lab = fields.IntField(null=True)
    # Связи с другими моделями
    control_works: fields.ReverseRelation["ControlWork"]
    literature: fields.ReverseRelation["Literature"]
    exam_questions: fields.ReverseRelation["ExamQuestion"]
    schedule_items: fields.ReverseRelation["ScheduleItem"]

    class Meta:
        table = "disciplines"

    def __str__(self):
        return self.name


class ControlWork(models.Model):
    """Модель контрольной работы"""
    id = fields.IntField(pk=True)
    number = fields.IntField(null=False)
    max_score = fields.IntField(null=True)
    week = fields.IntField(null=True)
    semester = fields.IntField(null=True)
    format = fields.CharField(max_length=50, null=True)  # очно, СДО
    discipline = fields.ForeignKeyField("models.Discipline", related_name="control_works", null=True)

    class Meta:
        table = "control_works"

    def __str__(self):
        return f"КР {self.number} - {self.discipline.name if self.discipline else ''}"


class Grade(models.Model):
    """Модель оценки/отметки"""
    id = fields.IntField(pk=True)
    student = fields.ForeignKeyField("models.Student", related_name="grades")
    control_work = fields.ForeignKeyField("models.ControlWork", related_name="grades")
    score = fields.IntField(null=True)
    date = fields.DateField(null=True)

    class Meta:
        table = "grades"
        unique_together = (("student", "control_work"),)

    def __str__(self):
        return f"{self.student.full_name if self.student else ''} - {self.score}"


class Literature(models.Model):
    """Модель литературы по дисциплине"""
    id = fields.IntField(pk=True)
    title = fields.CharField(max_length=500, null=False)
    authors = fields.CharField(max_length=500, null=True)
    publisher = fields.CharField(max_length=255, null=True)
    year = fields.IntField(null=True)
    discipline = fields.ForeignKeyField("models.Discipline", related_name="literature", null=True)

    class Meta:
        table = "literature"

    def __str__(self):
        return self.title


class ExamQuestion(models.Model):
    """Модель вопроса к экзамену/зачету"""
    id = fields.IntField(pk=True)
    number = fields.IntField(null=False)
    text = fields.TextField(null=False)
    discipline = fields.ForeignKeyField("models.Discipline", related_name="exam_questions", null=True)

    class Meta:
        table = "exam_questions"

    def __str__(self):
        return f"Вопрос {self.number}"


class Publication(models.Model):
    """Модель публикации студента"""
    id = fields.IntField(pk=True)
    title = fields.TextField(null=False)
    student = fields.ForeignKeyField("models.Student", related_name="publications", null=True)

    class Meta:
        table = "publications"

    def __str__(self):
        return self.title


class Template(models.Model):
    """Модель шаблона документа"""
    id = fields.IntField(pk=True)
    name = fields.CharField(max_length=255, null=False)
    file_path = fields.CharField(max_length=500, null=False)
    file_type = fields.CharField(max_length=10, null=False)  # docx, xlsx
    description = fields.TextField(null=True)

    class Meta:
        table = "templates"

    def __str__(self):
        return self.name


class TimeSlot(models.Model):
    """Модель временного слота для расписания"""
    id = fields.IntField(pk=True)
    number = fields.IntField(null=False)  # Номер пары
    start_time = fields.CharField(max_length=10, null=False)
    end_time = fields.CharField(max_length=10, null=False)

    class Meta:
        table = "time_slots"
        unique_together = (("number", "start_time", "end_time"),)

    def __str__(self):
        return f"Пара {self.number} ({self.start_time}-{self.end_time})"


class Classroom(models.Model):
    """Модель аудитории"""
    id = fields.IntField(pk=True)
    name = fields.CharField(max_length=50, null=False)
    capacity = fields.IntField(null=True)

    class Meta:
        table = "classrooms"

    def __str__(self):
        return self.name


class ScheduleItem(models.Model):
    """Модель элемента расписания"""
    id = fields.IntField(pk=True)
    day_of_week = fields.CharEnumField(DayOfWeek, null=False)  # День недели
    time_slot = fields.ForeignKeyField("models.TimeSlot", related_name="schedule_items")
    discipline = fields.ForeignKeyField("models.Discipline", related_name="schedule_items")
    teacher = fields.ForeignKeyField("models.Teacher", related_name="schedule_items")
    group = fields.ForeignKeyField("models.Group", related_name="schedule_items")
    classroom = fields.ForeignKeyField("models.Classroom", related_name="schedule_items", null=True)
    is_lecture = fields.BooleanField(default=False)  # Лекция или практика/лабораторная
    week_type = fields.CharField(max_length=20, null=True)  # Тип недели (числитель/знаменатель/еженедельно)

    class Meta:
        table = "schedule_items"
        unique_together = (("day_of_week", "time_slot", "group", "week_type"),)

    def __str__(self):
        return f"{self.day_of_week.value}, {self.time_slot}, {self.discipline.name}, {self.group.code}, {self.classroom.name if self.classroom else 'Нет аудитории'}"


@register(Teacher)
class TeacherAdmin(TortoiseModelAdmin):
    list_display = ("id", "full_name", "email")
    list_display_links = ("id", "full_name")
    search_fields = ("full_name", "email")


@register(Group)
class GroupAdmin(TortoiseModelAdmin):
    list_display = ("id", "code", "course", "teacher")
    list_display_links = ("id", "code")
    list_filter = ("teacher",)
    search_fields = ("code", "course")


@register(Student)
class StudentAdmin(TortoiseModelAdmin):
    list_display = ("id", "full_name", "email", "group")
    list_display_links = ("id", "full_name")
    list_filter = ("group",)
    search_fields = ("full_name", "email")


@register(ScheduleItem)
class ScheduleItemAdmin(TortoiseModelAdmin):
    pass


@register(Discipline)
class DisciplineAdmin(TortoiseModelAdmin):
    list_display = ("id", "name", "education_level", "department", "hours_lecture", "hours_practice", "hours_lab")
    list_display_links = ("id", "name")
    list_filter = ("education_level",)
    search_fields = ("name", "department")


@register(ControlWork)
class ControlWorkAdmin(TortoiseModelAdmin):
    list_display = ("id", "number", "discipline", "max_score", "week", "semester", "format")
    list_display_links = ("id",)
    list_filter = ("discipline", "format", "semester")


@register(Grade)
class GradeAdmin(TortoiseModelAdmin):
    list_display = ("id", "student", "control_work", "score", "date")
    list_display_links = ("id",)
    list_filter = ("student", "control_work", "date")


@register(Literature)
class LiteratureAdmin(TortoiseModelAdmin):
    list_display = ("id", "title", "authors", "publisher", "year", "discipline")
    list_display_links = ("id", "title")
    list_filter = ("discipline", "year")
    search_fields = ("title", "authors")


@register(ExamQuestion)
class ExamQuestionAdmin(TortoiseModelAdmin):
    list_display = ("id", "number", "discipline", "text")
    list_display_links = ("id", "number")
    list_filter = ("discipline",)
    search_fields = ("text",)


@register(Publication)
class PublicationAdmin(TortoiseModelAdmin):
    list_display = ("id", "title", "student")
    list_display_links = ("id",)
    list_filter = ("student",)
    search_fields = ("title",)


@register(Template)
class TemplateAdmin(TortoiseModelAdmin):
    list_display = ("id", "name", "file_path", "file_type", "description")
    list_display_links = ("id", "name")
    list_filter = ("file_type",)
    search_fields = ("name", "description")

    @action(description="Создать копию шаблона")
    async def duplicate_template(self, request, pk):
        template = await Template.get(id=pk)
        new_template = await Template.create(
            name=f"{template.name} (копия)",
            file_path=template.file_path,
            file_type=template.file_type,
            description=template.description
        )
        return True, f"Создана копия шаблона: {new_template.name}"


@register(AdminUser)
class UserAdmin(TortoiseModelAdmin):
    list_display = ("id", "username", "is_superuser", "is_active")
    list_display_links = ("id", "username")
    list_filter = ("id", "username", "is_superuser", "is_active")
    search_fields = ("username",)
    formfield_overrides = {
        "username": (WidgetType.SlugInput, {"required": True}),
        "password": (WidgetType.PasswordInput, {"passwordModalForm": True}),
    }
    actions = (
        *TortoiseModelAdmin.actions,
        "activate",
        "deactivate",
    )

    async def authenticate(self, username: str, password: str) -> int | None:
        user = await AdminUser.filter(username=username, is_superuser=True).first()
        if not user:
            return None
        if not bcrypt.checkpw(password.encode(), user.password.encode()):
            return None
        return user.pk

    async def change_password(self, id: int, password: str) -> None:
        user = await self.model_cls.filter(id=id).first()
        if not user:
            return

        user.set_password(password.encode())
        await user.save(update_fields=("password",))

    @action(description="Set as active")
    async def activate(self, ids: list[int]) -> None:
        await self.model_cls.filter(id__in=ids).update(is_active=True)

    @action(description="Deactivate")
    async def deactivate(self, ids: list[int]) -> None:
        await self.model_cls.filter(id__in=ids).update(is_active=False)
