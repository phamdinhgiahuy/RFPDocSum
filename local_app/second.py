from flask import Blueprint, render_template

second = Blueprint(
    "second", __name__, static_folder="static", template_folder="templates"
)


@second.route("/")
@second.route("/home/")
def home():
    return render_template("home.html")


@second.route("/about/")
def about():
    return render_template("about.html")
