from flask import Flask, render_template
from routes.labs import labs_bp
from routes.students import students_bp
from routes.subjects import subjects_bp
from routes.timetable import timetable_bp
from routes.step1 import step1_bp
from routes.staff import staff_bp

app = Flask(__name__)

app.register_blueprint(labs_bp)
app.register_blueprint(students_bp)
app.register_blueprint(subjects_bp)
app.register_blueprint(timetable_bp)
app.register_blueprint(step1_bp)
app.register_blueprint(staff_bp)
app.config['TEMPLATES_AUTO_RELOAD'] = True
@app.route("/")
def home():
    return render_template("home.html")

@app.route("/timetable_start")
def timetable_start():
    return render_template("index.html")

@app.route("/details")
def details():
    return render_template("details.html")

if __name__ == "__main__":
    app.run(debug=True)
