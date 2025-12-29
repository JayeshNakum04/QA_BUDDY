from flask import Flask, render_template, request, redirect, url_for
import sqlite3
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from flask import send_file
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'static/uploads'

def get_db():
    return sqlite3.connect('database.db')

@app.route('/')
def index():
    return render_template('index.html')

# ---------------- BUG SECTION ----------------

@app.route('/add-bug', methods=['GET', 'POST'])

def add_bug():
    if request.method == 'POST':
        title = request.form['title']
        steps = request.form['steps']
        expected = request.form['expected']
        actual = request.form['actual']
        severity = request.form['severity']
        priority = request.form['priority']


        screenshot = request.files.get('screenshot')
        filename = None

        if screenshot and screenshot.filename != "":
            filename = secure_filename(screenshot.filename)
            screenshot.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        db = get_db()
        db.execute("""
            INSERT INTO bugs (title, steps, expected, actual, severity, status, screenshot, priority)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (title, steps, expected, actual, severity, 'Open', filename, priority))


        db.commit()
        db.close()

        return redirect('/bugs')

    return render_template('add_bug.html')


@app.route('/bugs')
def bugs():
    status = request.args.get('status')
    priority = request.args.get('priority')

    query = "SELECT * FROM bugs WHERE 1=1"
    params = []

    if status:
        query += " AND status = ?"
        params.append(status)

    if priority:
        query += " AND priority = ?"
        params.append(priority)

    db = get_db()
    bugs = db.execute(query, params).fetchall()
    db.close()

    return render_template('bugs.html', bugs=bugs)


@app.route('/bug/<int:bug_id>')
def bug_detail(bug_id):
    db = get_db()
    bug = db.execute(
        "SELECT * FROM bugs WHERE id = ?", (bug_id,)
    ).fetchone()
    db.close()

    return render_template('bug_detail.html', bug=bug)

@app.route('/update-status/<int:bug_id>', methods=['POST'])
def update_status(bug_id):
    new_status = request.form['status']

    db = get_db()
    db.execute(
        "UPDATE bugs SET status = ? WHERE id = ?",
        (new_status, bug_id)
    )
    db.commit()
    db.close()

    return redirect(f'/bug/{bug_id}')


@app.route('/delete-bug/<int:bug_id>')
def delete_bug(bug_id):
    db = get_db()
    db.execute("DELETE FROM bugs WHERE id = ?", (bug_id,))
    db.commit()
    db.close()
    return redirect('/bugs')


# ---------------- TEST CASE SECTION ----------------

@app.route('/testcase/<int:tc_id>')
def testcase_detail(tc_id):
    db = get_db()
    tc = db.execute(
        "SELECT * FROM testcases WHERE id = ?", (tc_id,)
    ).fetchone()
    db.close()

    return render_template('testcase_detail.html', tc=tc)


@app.route('/add-testcase', methods=['GET', 'POST'])
def add_testcase():
    if request.method == 'POST':
        module = request.form['module']
        scenario = request.form['scenario']
        steps = request.form['steps']
        expected = request.form['expected']
        status = request.form['status']

        db = get_db()
        db.execute("""
            INSERT INTO testcases (module, scenario, steps, expected, status)
            VALUES (?, ?, ?, ?, ?)
        """, (module, scenario, steps, expected, status))
        db.commit()
        db.close()

        return redirect('/testcases')

    return render_template('add_testcase.html')


@app.route('/testcases')
def testcases():
    status = request.args.get('status')

    query = "SELECT * FROM testcases WHERE 1=1"
    params = []

    if status:
        query += " AND status = ?"
        params.append(status)

    db = get_db()
    testcases = db.execute(query, params).fetchall()
    db.close()

    return render_template('testcases.html', testcases=testcases)


@app.route('/update-testcase-status/<int:tc_id>', methods=['POST'])
def update_testcase_status(tc_id):
    new_status = request.form['status']

    db = get_db()
    db.execute(
        "UPDATE testcases SET status = ? WHERE id = ?",
        (new_status, tc_id)
    )
    db.commit()
    db.close()

    return redirect('/testcases')

@app.route('/delete-testcase/<int:tc_id>')
def delete_testcase(tc_id):
    db = get_db()
    db.execute("DELETE FROM testcases WHERE id = ?", (tc_id,))
    db.commit()
    db.close()
    return redirect('/testcases')

# ---------------- EXPORT SECTION ----------------

@app.route('/export-bugs')
def export_bugs():
    db = get_db()
    bugs = db.execute("SELECT * FROM bugs").fetchall()
    db.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Bugs Report"

    # Header
    ws.append([
        "S.No", "Title", "Severity", "Priority",
        "Status", "Steps", "Expected Result", "Actual Result"
    ])

    # Data
    for i, bug in enumerate(bugs, start=1):
        ws.append([
            i,
            bug[1],   # title
            bug[5],   # severity
            bug[8],   # priority
            bug[6],   # status
            bug[2],   # steps
            bug[3],   # expected
            bug[4]    # actual
        ])

    file_path = "bugs_report.xlsx"
    wb.save(file_path)

    return send_file(file_path, as_attachment=True)

@app.route('/export-testcases')
def export_testcases():
    db = get_db()
    testcases = db.execute("SELECT * FROM testcases").fetchall()
    db.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Test Cases"

    # Header
    ws.append([
        "S.No", "Module", "Scenario",
        "Steps", "Expected Result", "Status"
    ])

    # Data
    for i, tc in enumerate(testcases, start=1):
        ws.append([
            i,
            tc[1],  # module
            tc[2],  # scenario
            tc[3],  # steps
            tc[4],  # expected
            tc[5]   # status
        ])

    file_path = "testcases_report.xlsx"
    wb.save(file_path)

    return send_file(file_path, as_attachment=True)



if __name__ == '__main__':
    app.run(debug=True)
