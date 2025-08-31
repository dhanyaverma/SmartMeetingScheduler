import pandas as pd
import datetime
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def time_to_minutes(t):
    if pd.isna(t):
        return None
    if isinstance(t, pd.Timestamp):
        return t.hour * 60 + t.minute
    if isinstance(t, datetime.time):
        return t.hour * 60 + t.minute
    t_str = str(t).strip()
    if t_str == "":
        return None
    for fmt in ("%I:%M %p", "%I:%M:%S %p", "%H:%M", "%H:%M:%S"):
        try:
            dt = datetime.datetime.strptime(t_str, fmt)
            return dt.hour * 60 + dt.minute
        except ValueError:
            continue
    try:
        f = float(t)
        minutes = int(f * 24 * 60)
        return minutes
    except:
        pass
    logger.warning(f"Cannot parse time: {t}")
    return None

def build_busy_map(df):
    expected_columns = ["START TIME", "END TIME", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"]
    if not all(col in df.columns for col in expected_columns):
        raise ValueError(f"Excel sheet missing required columns: {expected_columns}")
    
    busy = {day.upper(): [] for day in expected_columns[2:]}
    for idx, row in df.iterrows():
        start = time_to_minutes(row["START TIME"])
        end = time_to_minutes(row["END TIME"])
        if start is None or end is None or end <= start:
            logger.warning(f"Skipping invalid row {idx}: start={row['START TIME']}, end={row['END TIME']}")
            continue
        for day in busy.keys():
            lecture = row[day]
            if pd.notna(lecture) and str(lecture).strip() != "":
                busy[day].append((start, end))
    return busy