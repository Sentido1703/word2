from flask import Blueprint, request
from services.report_service import generate_report

report_bp = Blueprint('report_bp', __name__)


@report_bp.route("/officesvc/report", methods=["GET"])
def create_report():
    report_type = request.args.get("type")
    group_id = request.args.get("id")
    template_name = request.args.get("template")

    if not report_type or not group_id or not template_name:
        return {"error": "Не хватает параметров 'type', 'id' или 'template'"}, 400

    return generate_report(report_type, group_id, template_name)
