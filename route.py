from flask import Blueprint, render_template, session, redirect
from Execute.Functions import functions

routes_bp = Blueprint("routes_bp", __name__)


def get_user_context():
    user = session.get("user", {})
    return {
        "user":          user,
        "user_location": user.get("location", ""),
        "user_role":     user.get("role", "user")
    }


# =====================================================
# MAIN PAGE
# =====================================================

@routes_bp.route("/")
def home():
    return render_template("index.html", **get_user_context())


# =====================================================
# ATTENDANCE REPORT APIs
# =====================================================

routes_bp.add_url_rule(
    "/generate_report",
    view_func=functions.generate_report_fn,
    methods=["POST"]
)

routes_bp.add_url_rule(
    "/get_locations",
    view_func=functions.get_locations_fn,
    methods=["GET"]
)

routes_bp.add_url_rule(
    "/get_months",
    view_func=functions.get_months_fn,
    methods=["GET"]
)

routes_bp.add_url_rule(
    "/get_organizations",
    view_func=functions.get_organizations_by_filter_fn,
    methods=["GET"]
)

routes_bp.add_url_rule(
    "/download_excel",
    view_func=functions.download_excel_fn,
    methods=["GET"]
)
