import os
from docx import Document
from flask import jsonify, current_app
from models.models import Group


def generate_report(report_type, group_id, template_name):
    TEMPLATE_DIR = current_app.config['TEMPLATE_DIR']
    OUTPUT_DIR = current_app.config['OUTPUT_DIR']

    group = Group.query.filter_by(id=group_id).first()
    if not group:
        return jsonify({"error": "Группа с данным ID не найдена"}), 404

    template_path = os.path.join(TEMPLATE_DIR, f"{template_name}.docx")
    if not os.path.exists(template_path):
        return jsonify({"error": f"Шаблон {template_name}.docx не найден"}), 404

    document = Document(template_path)
    for paragraph in document.paragraphs:
        if "{{course}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{course}}", group.course)
        if "{{students}}" in paragraph.text:
            students_text = "\n".join(
                [f"{student.name} ({student.email})" for student in group.students]
            )
            paragraph.text = paragraph.text.replace("{{students}}", students_text)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, f"{report_type}_{group_id}.docx")
    document.save(output_path)

    return jsonify({
        "message": "Отчет успешно сгенерирован и сохранён!",
        "file_path": output_path
    }), 200
