from flask import Flask, render_template, request, session, redirect, url_for
import pandas as pd
import os
from smart_scheduler import time_to_minutes, build_busy_map, expected_columns
from datetime import datetime, date
from collections import defaultdict

app = Flask(__name__)
app.secret_key = 'super_secret_key'  # Change in production

# Excel path
excel_path = os.path.join(os.path.dirname(__file__), "faculty_timetable.xlsx")

# In-memory temporary changes: {(faculty, date_str): DataFrame}
temp_sheets = {}

# Load original timetables
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

# Time formatting
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
        minutes = time_to_minutes(t)
        if minutes is None:
            return str(t)
        return to_hhmm(minutes)
    except Exception as e:
        print(f"Error in to_12hour: {e}")
        return str(t)

# Register Jinja2 filter
app.jinja_env.filters['to_12hour'] = to_12hour
print("Registered to_12hour filter")

def join_time(hour, minute, ampm):
    return f"{hour}:{minute} {ampm}"

def clear_expired_temp_sheets():
    """Remove temporary edits for past dates."""
    current_date = date.today()
    keys_to_remove = [
        key for key, df in temp_sheets.items()
        if datetime.strptime(key[1], '%Y-%m-%d').date() < current_date
    ]
    for key in keys_to_remove:
        del temp_sheets[key]

@app.route("/", methods=["GET", "POST"])
def index():
    clear_expired_temp_sheets()
    teachers = list(sheets.keys())
    today = datetime.now().strftime('%Y-%m-%d')
    if not teachers:
        return render_template("index.html", error="Error: Could not load faculty timetables.", teachers=[], today=today)

    if request.method == "POST":
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
    date_str = session['form_data']['date'][0]
    if faculty not in sheets:
        return "Faculty not found", 404
    df = temp_sheets.get((faculty, date_str), sheets[faculty])
    df = df.fillna('')  # Remove NaN
    columns = df.columns.tolist()
    rows = df.to_dict('records')
    return render_template("edit_timetable.html", faculty=faculty, columns=columns, rows=rows)

@app.route("/save_edit/<faculty>", methods=["POST"])
def save_edit(faculty):
    if faculty not in sheets:
        return "Faculty not found", 404
    date_str = session['form_data']['date'][0]
    columns = expected_columns
    data = defaultdict(dict)
    deleted = set()

    for key in request.form:
        if key.startswith('delete_'):
            del_idx = key.split('_', 1)[1]
            deleted.add(del_idx)

    for key, value in request.form.items():
        if '_' in key and not key.startswith('delete_'):
            col, idx = key.rsplit('_', 1)
            data[idx][col] = value

    rows = []
    for idx, row_data in data.items():
        if idx in deleted:
            continue
        row = {col: row_data.get(col, '') for col in columns}
        rows.append(row)

    new_df = pd.DataFrame(rows, columns=columns)
    temp_sheets[(faculty, date_str)] = new_df

    return render_template("select_faculty.html", teachers=session['selected_teachers'])

@app.route("/process")
def process():
    clear_expired_temp_sheets()
    form_data = session.get('form_data')
    if not form_data:
        return redirect(url_for("index"))

    result_list = []
    alternative_slots = []
    error = None
    day_display = None
    selected_teachers = form_data['teacher']

    try:
        if not selected_teachers:
            raise ValueError("Please select at least one faculty.")
        
        duration = int(form_data['duration'][0])
        if duration <= 0:
            raise ValueError("Duration must be a positive number.")

        # Parse date
        date_str = form_data['date'][0]
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        day = date_obj.strftime('%A').upper()
        day_display = date_obj.strftime('%B %d, %Y')
        if day == 'SUNDAY':
            raise ValueError("No timetable available for Sundays.")

        # Parse time window
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

        ws_min = time_to_minutes(window_start)
        we_min = time_to_minutes(window_end)
        if ws_min is None or we_min is None or we_min <= ws_min:
            raise ValueError("Invalid time window (start must be before end).")

        # Build busy maps
        busy_map = {}
        for t in selected_teachers:
            df = temp_sheets.get((t, date_str), sheets[t])
            busy_map[t] = build_busy_map(df)

        # === 1. Search in preferred window ===
        for start in range(ws_min, we_min - duration + 1, 15):
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

        # === 2. If no slots, search extended window (±2 hours) ===
        if not result_list:
            EXTEND_MINUTES = 120
            search_start = max(0, ws_min - EXTEND_MINUTES)
            search_end = min(24*60, we_min + EXTEND_MINUTES)

            candidates = []
            for start in range(search_start, search_end - duration + 1, 15):
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
                    center_user = (ws_min + we_min) / 2
                    center_slot = (start + end) / 2
                    distance = abs(center_slot - center_user)
                    candidates.append((distance, start, end))

            # Sort by closeness → pick top 5
            candidates.sort(key=lambda x: x[0])
            closest_5 = [(s, e) for _, s, e in candidates[:5]]

            # Sort by start time for display
            closest_5.sort(key=lambda x: x[0])

            alternative_slots = closest_5

        # Format results
        result_list = [{"Start": to_hhmm(s), "End": to_hhmm(e)} for s, e in result_list]
        alternative_slots = [{"Start": to_hhmm(s), "End": to_hhmm(e)} for s, e in alternative_slots]

    except Exception as e:
        error = f"Error processing request: {str(e)}"

    session.pop('form_data', None)
    session.pop('selected_teachers', None)

    return render_template(
        "result.html",
        results=result_list,
        alternatives=alternative_slots,
        teachers=", ".join(selected_teachers),
        day=day_display,
        error=error
    )

if __name__ == "__main__":
    app.run(debug=True)