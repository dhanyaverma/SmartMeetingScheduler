from flask import Flask, render_template, request
import pandas as pd
import os
from smart_scheduler import time_to_minutes, build_busy_map
from datetime import datetime

app = Flask(__name__)

# Load Excel file with error handling
try:
    excel_path = os.path.join(os.path.dirname(__file__), "faculty_timetable.xlsx")
    sheets = pd.read_excel(excel_path, sheet_name=None)
    if not sheets:
        raise ValueError("Excel file is empty or has no valid sheets.")
except Exception as e:
    print(f"Error loading Excel file: {e}")
    sheets = {}

def to_hhmm(minutes):
    if minutes is None:
        return "Invalid"
    h = minutes // 60
    m = minutes % 60
    ampm = "AM" if h < 12 else "PM"
    h = h if 1 <= h <= 12 else (h - 12 if h > 12 else 12)
    return f"{h}:{m:02d} {ampm}"

def join_time(hour, minute, ampm):
    return f"{hour}:{minute} {ampm}"

@app.route("/", methods=["GET", "POST"])
def index():
    teachers = list(sheets.keys())  # Dynamically get all sheet names as faculties
    today = datetime.now().strftime('%Y-%m-%d')  # Pass today's date to template
    if not teachers:
        return render_template("index.html", error="Error: Could not load faculty timetables.", teachers=[], today=today)

    result_list = []
    error = None
    day_display = None
    selected_teachers = []

    if request.method == "POST":
        try:
            selected_teachers = request.form.getlist("teacher")
            if not selected_teachers:
                raise ValueError("Please select at least one faculty.")
            
            duration = int(request.form.get("duration"))
            if duration <= 0:
                raise ValueError("Duration must be a positive number.")

            # Get date and convert to weekday
            date_str = request.form.get("date")
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
                request.form.get("window_start_hour"),
                request.form.get("window_start_minute"),
                request.form.get("window_start_ampm")
            )
            window_end = join_time(
                request.form.get("window_end_hour"),
                request.form.get("window_end_minute"),
                request.form.get("window_end_ampm")
            )

            # Validate window
            ws_min = time_to_minutes(window_start)
            we_min = time_to_minutes(window_end)
            if ws_min is None or we_min is None or we_min <= ws_min:
                raise ValueError("Invalid time window (start must be before end).")

            # Build busy maps for selected teachers
            busy_map = {t: build_busy_map(sheets[t]) for t in selected_teachers}

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

        return render_template(
            "result.html",
            results=result_list,
            teachers=", ".join(selected_teachers),
            day=day_display,
            error=error
        )

    return render_template("index.html", teachers=teachers, error=None, today=today)

if __name__ == "__main__":
    app.run(debug=True)