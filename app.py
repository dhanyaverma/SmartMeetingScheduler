from flask import Flask, render_template, request, session, redirect, url_for
import pandas as pd
import os
from smart_scheduler import time_to_minutes, build_busy_map, expected_columns
from datetime import datetime, date
from collections import defaultdict

app = Flask(__name__)
app.secret_key = 'super_secret_key'  # Required for sessions; change in production

# Excel path
excel_path = os.path.join(os.path.dirname(__file__), "faculty_timetable.xlsx")

# Store temporary timetable changes: {(faculty, date_str): DataFrame}
temp_sheets = {}

# Function to load/reload sheets
def load_sheets():
    try:
        sheets = pd.read_excel(excel_path, sheet_name=None)
        if not sheets:
            raise ValueError("Excel file is empty or has no valid sheets.")
        return sheets
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return {}

sheets = load_sheets()

def to_hhmm(minutes):
    if minutes is None:
        return "Invalid"
    h = minutes // 60
    m = minutes % 60
    ampm = "AM" if h < 12 else "PM"
    h = h if 1 <= h <= 12 else (h - 12 if h > 12 else 12)
    return f"{h}:{m:02d} {ampm}"

def to_12hour(t):
    """Convert time string to 12-hour format, handle NaN/None."""
    if pd.isna(t) or t is None or str(t).strip() == "":
        return ""
    try:
        # Parse time string using time_to_minutes for consistency
        minutes = time_to_minutes(t)
        if minutes is None:
            return str(t)  # Return as-is if unparseable
        return to_hhmm(minutes)
    except Exception as e:
        print(f"Error in to_12hour: {e}")
        return str(t)  # Fallback to original string if parsing fails

# Register custom Jinja2 filter
app.jinja_env.filters['to_12hour'] = to_12hour
print("Registered to_12hour filter")  # Debug to confirm registration

def join_time(hour, minute, ampm):
    return f"{hour}:{minute} {ampm}"

def clear_expired_temp_sheets():
    """Clear temporary sheets for dates that have passed."""
    current_date = date.today()
    keys_to_remove = [
        key for key, df in temp_sheets.items()
        if datetime.strptime(key[1], '%Y-%m-%d').date() < current_date
    ]
    for key in keys_to_remove:
        del temp_sheets[key]

@app.route("/", methods=["GET", "POST"])
def index():
    clear_expired_temp_sheets()  # Clear expired changes on each request
    teachers = list(sheets.keys())  # Dynamically get all sheet names as faculties
    today = datetime.now().strftime('%Y-%m-%d')  # Pass today's date to template
    if not teachers:
        return render_template("index.html", error="Error: Could not load faculty timetables.", teachers=[], today=today)

    if request.method == "POST":
        # Store form data in session (flat=False to preserve lists)
        session['form_data'] = request.form.to_dict(flat=False)
        return render_template("edit_prompt.html")

    return render_template("index.html", teachers=teachers, error=None, today=today)

@app.route("/handle_edit_prompt", methods=["POST"])
def handle_edit_prompt():
    choice = request.form.get("choice")
    if choice == "no":
        return redirect(url_for("process"))
    else:
        session['selected_teachers'] = session['form_data']['teacher']
        return render_template("select_faculty.html", teachers=session['selected_teachers'])

@app.route("/edit", methods=["POST"])
def edit():
    faculty = request.form.get("faculty")
    date_str = session['form_data']['date'][0]  # Get selected date
    if faculty not in sheets:
        return "Faculty not found", 404
    # Use temporary sheet if available for this faculty and date, else original
    df = temp_sheets.get((faculty, date_str), sheets[faculty])
    # Replace NaN with empty strings
    df = df.fillna('')
    columns = df.columns.tolist()
    rows = df.to_dict('records')
    return render_template("edit_timetable.html", faculty=faculty, columns=columns, rows=rows)

@app.route("/save_edit/<faculty>", methods=["POST"])
def save_edit(faculty):
    if faculty not in sheets:
        return "Faculty not found", 404
    date_str = session['form_data']['date'][0]  # Get selected date
    columns = expected_columns
    data = defaultdict(dict)
    deleted = set()

    # Collect deleted indices
    for key in request.form:
        if key.startswith('delete_'):
            del_idx = key.split('_', 1)[1]
            deleted.add(del_idx)

    # Collect data from form
    for key, value in request.form.items():
        if '_' in key and not key.startswith('delete_'):
            col, idx = key.rsplit('_', 1)
            data[idx][col] = value

    # Build new rows
    rows = []
    for idx, row_data in data.items():
        if idx in deleted:
            continue
        row = {col: row_data.get(col, '') for col in columns}
        rows.append(row)

    # Store in temp_sheets instead of writing to Excel
    new_df = pd.DataFrame(rows, columns=columns)
    temp_sheets[(faculty, date_str)] = new_df

    # Return to select faculty for more edits
    return render_template("select_faculty.html", teachers=session['selected_teachers'])

@app.route("/process")
def process():
    clear_expired_temp_sheets()  # Clear expired changes
    form_data = session.get('form_data')
    if not form_data:
        return redirect(url_for("index"))

    result_list = []
    error = None
    day_display = None
    selected_teachers = form_data['teacher']

    try:
        if not selected_teachers:
            raise ValueError("Please select at least one faculty.")
        
        duration = int(form_data['duration'][0])
        if duration <= 0:
            raise ValueError("Duration must be a positive number.")

        # Get date and convert to weekday
        date_str = form_data['date'][0]
        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            day = date_obj.strftime('%A').upper()  # e.g., 'FRIDAY'
            day_display = date_obj.strftime('%B %d, %Y')  # e.g., 'August 29, 2025'
            if day == 'SUNDAY':
                raise ValueError("No timetable available for Sundays.")
        except ValueError as e:
            raise ValueError(f"Invalid date: {str(e)}")

        # Window times
        window_start = join_time(
            form_data['window_start_hour'][0],
            form_data['window_start_minute'][0],
            form_data['window_start_ampm'][0]
        )
        window_end = join_time(
            form_data['window_end_hour'][0],
            form_data['window_end_minute'][0],
            form_data['window_end_ampm'][0]
        )

        # Validate window
        ws_min = time_to_minutes(window_start)
        we_min = time_to_minutes(window_end)
        if ws_min is None or we_min is None or we_min <= ws_min:
            raise ValueError("Invalid time window (start must be before end).")

        # Build busy maps for selected teachers
        busy_map = {}
        for t in selected_teachers:
            # Use temporary sheet for this faculty and date, else original
            df = temp_sheets.get((t, date_str), sheets[t])
            busy_map[t] = build_busy_map(df)

        # Search window for available slots
        for start in range(ws_min, we_min - duration + 1, 15):  # Step by 15 min
            end = start + duration
            conflict = False
            for t in selected_teachers:
                for s, e in busy_map[t].get(day, []):
                    if not (end <= s or start >= e):
                        conflict = True
                        break
                if conflict:
                    break
            if not conflict:
                result_list.append((start, end))

        # Format results
        result_list = [{"Start": to_hhmm(s), "End": to_hhmm(e)} for s, e in result_list]

    except Exception as e:
        error = f"Error processing request: {str(e)}"

    # Clear session after processing
    session.pop('form_data', None)
    session.pop('selected_teachers', None)

    return render_template(
        "result.html",
        results=result_list,
        teachers=", ".join(selected_teachers),
        day=day_display,
        error=error
    )

if __name__ == "__main__":
    app.run(debug=True)