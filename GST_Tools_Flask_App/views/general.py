from flask import Blueprint, render_template

general_views_bp = Blueprint(__name__, "general_view_bp")


@general_views_bp.route("/")
def home_view():
    return render_template("general/home.html")
