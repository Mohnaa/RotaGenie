# rota_month_streamlit.py
import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
from calendar import monthrange

st.set_page_config(page_title="Monthly Rota Generator", layout="wide")

st.title("Monthly Rota Generator — Calendar layout + Leave support")

st.markdown("""
**How to use**
- Enter team members (comma separated). Make sure `Priyam` is included (he defaults to UK).
- Enter manual leaves lines (one per line): `Name dd-mm-yyyy:dd-mm-yyyy`
- Choose month & year, and mode (Mon–Fri or Mon–Sun).
- The app uses your 5-week rotation mapping by default; you can edit the sequences if needed.
""")

# ---------------- Inputs ----------------
with st.sidebar:
    st.header("Team & month")
    team_input = st.text_area("Team members (comma separated)",
                              value="Mohnaa, Vamsi, Usha, Simran, Priyam, Abhay, Sanskar",
                              height=120)
    members = [m.strip() for m in team_input.split(",") if m.strip()]

    month = st.selectbox("Month", list(range(1,13)), index=datetime.today().month-1)
    year = st.number_input("Year", min_value=2000, max_value=2100, value=datetime.today().year)

    st.markdown("---")
    st.header("Working days mode")
    mode = st.radio("Choose working days for each week:",
                    ("Mon–Fri (Normal)", "Mon–Sun (Festive)"))
    seven_day_mode = (mode == "Mon–Sun (Festive)")

    st.markdown("---")
    st.header("Leaves (manual)")
    st.markdown("Format: `Name dd-mm-yyyy:dd-mm-yyyy` (one per line). Single-day leave: use same date twice.")
    leave_text = st.text_area("Manual leaves", height=140,
                              value="Usha 24-11-2025:28-11-2025\nMohnaa 28-11-2025:28-11-2025")

    st.markdown("---")
    st.header("Rotation sequences (editable)")
    st.markdown("The app uses these sequences to pick Morning/General/Night across weeks (week index → pick element).")
    # Default sequences derived from your provided mapping (5 entries)
    # Morning: Usha, Simran, Sanskar, Mohnaa, Vamsi
    # General: Simran, Sanskar, Mohnaa, Vamsi, Abhay
    # Night: Sanskar, Mohnaa, Vamsi, Abhay, Usha
    morning_seq_text = st.text_input("Morning sequence (comma separated)",
                                     value="Usha,Simran,Sanskar,Mohnaa,Vamsi")
    general_seq_text = st.text_input("General sequence (comma separated)",
                                     value="Simran,Sanskar,Mohnaa,Vamsi,Abhay")
    night_seq_text = st.text_input("Night sequence (comma separated)",
                                   value="Sanskar,Mohnaa,Vamsi,Abhay,Usha")

    st.markdown("---")
    st.write("Extras")
    download_excel = st.checkbox("Enable Excel (.xlsx) download", value=True)

# ---------------- Parse rotation sequences ----------------
def parse_seq(s):
    return [x.strip() for x in s.split(",") if x.strip()]

morning_seq = parse_seq(morning_seq_text)
general_seq = parse_seq(general_seq_text)
night_seq = parse_seq(night_seq_text)

# ---------------- Parse leaves ----------------
def parse_leave_lines(text):
    leave_map = {}  # name -> set(date)
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        # allow "Name dd-mm-yyyy:dd-mm-yyyy"
        parts = line.split()
        if len(parts) < 2:
            st.warning(f"Bad leave line (ignored): {line}")
            continue
        name = parts[0].strip()
        dates_part = parts[1].strip()
        if ":" in dates_part:
            sdate_str, edate_str = dates_part.split(":",1)
        else:
            sdate_str = edate_str = dates_part
        try:
            sdt = datetime.strptime(sdate_str.strip(), "%d-%m-%Y").date()
            edt = datetime.strptime(edate_str.strip(), "%d-%m-%Y").date()
        except Exception as e:
            st.warning(f"Couldn't parse dates in line: {line}")
            continue
        if name not in leave_map:
            leave_map[name] = set()
        cur = sdt
        while cur <= edt:
            leave_map[name].add(cur)
            cur = cur + timedelta(days=1)
    return leave_map

leave_map = parse_leave_lines(leave_text)

# ---------------- Build weeks for the month ----------------
def month_weeks(year, month, seven_day_mode):
    # return list of week_start dates (Monday) that overlap the month
    first_day = date(year, month, 1)
    last_day = date(year, month, monthrange(year, month)[1])
    # find the first Monday on or before first_day
    first_monday = first_day - timedelta(days=first_day.weekday())
    weeks = []
    cur = first_monday
    while cur <= last_day:
        weeks.append(cur)
        cur += timedelta(days=7)
    # filter weeks that have at least one day in the month and at least one working day depending on mode
    valid_weeks = []
    for w in weeks:
        # collect week dates based on mode (Mon-Sun or Mon-Fri)
        days = []
        if seven_day_mode:
            day_count = 7
        else:
            day_count = 5
        for i in range(day_count):
            day = w + timedelta(days=i)
            if first_day <= day <= last_day:
                days.append(day)
        if days:
            valid_weeks.append(w)
    return valid_weeks

weeks = month_weeks(year, month, seven_day_mode)

# ---------------- Helper: check availability for a week ----------------
def person_available_for_week(person, week_start, seven_day_mode):
    """
    Returns:
      - 'full_off' if person is on leave for all working days of that week
      - 'partial' if person has some leave on working days but not all
      - 'ok' if no leave on working days
    """
    if seven_day_mode:
        day_count = 7
    else:
        day_count = 5
    working_days = [week_start + timedelta(days=i) for i in range(day_count)]
    # filter days within the month? We treat days outside month as non-working for this context.
    # But availability should consider only the working_days used for assignment (we will only assign days that fall in month)
    dates_considered = working_days
    leave_dates = leave_map.get(person, set())
    leave_on = [d for d in dates_considered if d in leave_dates]
    if len(leave_on) == 0:
        return "ok"
    if len(leave_on) == len(dates_considered):
        return "full_off"
    return "partial"

# ---------------- Choose person for role, given sequence and week index ----------------
def choose_person_for_role(seq, week_idx, members_allowed, week_start, seven_day_mode):
    """
    seq: list of names (rotation sequence, length 5 assumed)
    week_idx: 0-based index of current week in the month (we map to week number modulo 5)
    members_allowed: list of candidates (e.g., all except Priyam)
    returns: chosen person name or 'UNAVAILABLE'
    """
    if not seq:
        return "UNAVAILABLE"
    # map week number to index in seq. We'll use week_idx % len(seq)
    start_index = week_idx % len(seq)
    # rotate through seq starting at start_index, but only choose those who are in members_allowed
    n = len(seq)
    for offset in range(n):
        cand = seq[(start_index + offset) % n]
        if cand not in members_allowed:
            continue
        av_status = person_available_for_week(cand, week_start, seven_day_mode)
        # rule: if cand has any leave on working days, they cannot be assigned to Morning/General/Night
        if av_status == "ok":
            return cand
    # If no fully available candidate found, try partial: allow partial only if that person still has at least one free working day (we'll assign them UK instead normally)
    for offset in range(n):
        cand = seq[(start_index + offset) % n]
        if cand not in members_allowed:
            continue
        av_status = person_available_for_week(cand, week_start, seven_day_mode)
        if av_status == "partial":
            # do not assign non-UK role if partial leave (per your rule). So skip.
            continue
    return "UNAVAILABLE"

# ---------------- Build calendar columns (all working dates in month) ----------------
def all_working_dates_in_month(year, month, seven_day_mode):
    first_day = date(year, month, 1)
    last_day = date(year, month, monthrange(year, month)[1])
    days = []
    cur = first_day
    while cur <= last_day:
        if seven_day_mode:
            days.append(cur)
        else:
            # only Mon-Fri
            if cur.weekday() < 5:
                days.append(cur)
        cur += timedelta(days=1)
    return days

working_dates = all_working_dates_in_month(year, month, seven_day_mode)

# ---------------- Main assignment per week ----------------
# Pre-calc week -> role assignments
week_role_assignment = {}  # week_start -> { 'Morning': name, 'General': name, 'Night': name }
members_no_priyam = [m for m in members if m.lower() != "priyam".lower()]

for idx, wstart in enumerate(weeks):
    # determine which seq index to use: we want Week1->index0, Week2->1, etc.
    # idx already gives week order in the month (0-based).
    # Choose for each role:
    morning_person = choose_person_for_role(morning_seq, idx, members_no_priyam, wstart, seven_day_mode)
    general_person = choose_person_for_role(general_seq, idx, members_no_priyam, wstart, seven_day_mode)
    night_person = choose_person_for_role(night_seq, idx, members_no_priyam, wstart, seven_day_mode)

    # ensure no duplicate among the three roles: if a person was chosen twice, try to pick next available for the second role
    # simple resolution: if duplicates exist, try to replace duplicate in order General then Night using next available in sequence
    def resolve_dup(primary, secondary_person, seq, week_idx):
        if secondary_person == primary:
            # find next available candidate in seq (excluding 'primary')
            n = len(seq)
            for offset in range(1, n):
                candidate = seq[(week_idx + offset) % n]
                if candidate == primary:
                    continue
                if candidate not in members_no_priyam:
                    continue
                if person_available_for_week(candidate, wstart, seven_day_mode) == "ok":
                    return candidate
            return "UNAVAILABLE"
        return secondary_person

    general_person = resolve_dup(morning_person, general_person, general_seq, idx)
    night_person = resolve_dup(morning_person, night_person, night_seq, idx)
    # ensure general and night different
    if general_person == night_person:
        night_person = resolve_dup(general_person, night_person, night_seq, idx)

    week_role_assignment[wstart] = {
        "Morning": morning_person,
        "General": general_person,
        "Night": night_person
    }

# ---------------- Build per-person per-date table ----------------
# initialize DataFrame rows for each person
cols = []
# header rows will be date and day
cols = working_dates  # list of date objects

# prepare a mapping date -> week_start it belongs to (Monday of its week)
def date_to_week_start(d):
    return d - timedelta(days=d.weekday())

date_week_map = {d: date_to_week_start(d) for d in cols}

# Prepare data: rows per person -> columns per date
table = []
for person in members:
    row = {"Name": person}
    for d in cols:
        wstart = date_week_map[d]
        # if week not in our weeks list (rare), mark UNAVAILABLE
        if wstart not in week_role_assignment:
            row[d] = "UNAVAILABLE"
            continue

        # Priyam default UK unless full_off
        if person.lower() == "priyam".lower():
            # if Priyam is full_off for that week -> UNAVAILABLE, else UK
            avail = person_available_for_week(person, wstart, seven_day_mode)
            if avail == "full_off":
                row[d] = "UNAVAILABLE"
            else:
                row[d] = "UK"
            continue

        # who is assigned to roles for this week?
        roles = week_role_assignment[wstart]
        # if person is the morning_person and week status ok -> Morning
        if roles["Morning"] == person:
            # double-check availability: if person is 'ok' -> Morning, if partial or full_off -> UNAVAILABLE
            av = person_available_for_week(person, wstart, seven_day_mode)
            if av == "ok":
                row[d] = "Morning"
            else:
                row[d] = "UNAVAILABLE"
            continue
        if roles["General"] == person:
            av = person_available_for_week(person, wstart, seven_day_mode)
            if av == "ok":
                row[d] = "General"
            else:
                row[d] = "UNAVAILABLE"
            continue
        if roles["Night"] == person:
            av = person_available_for_week(person, wstart, seven_day_mode)
            if av == "ok":
                row[d] = "Night"
            else:
                row[d] = "UNAVAILABLE"
            continue

        # otherwise person is UK by default for that week IF they have at least one working day available
        av = person_available_for_week(person, wstart, seven_day_mode)
        if av == "full_off":
            row[d] = "UNAVAILABLE"
        else:
            row[d] = "UK"
    table.append(row)

# Build DataFrame with columns as strings "Day\nDD"
col_labels = []
for d in cols:
    col_labels.append(f"{d.strftime('%a')}\n{d.day:02d}")

df = pd.DataFrame(table)
# reorder columns: Name first, then date columns in order
ordered_cols = ["Name"] + cols
df = df.set_index("Name")[cols].reset_index()
# df currently has date objects as columns; convert them to strings for display
df_display = df.copy()
df_display.columns = ["Name"] + [f"{d.strftime('%a')}\n{d.day:02d}" for d in cols]

# ---------------- Reorder rows: move Night shift people to bottom (as requested) ----------------
# We'll compute each person's primary role for the first week they appear, then sort
def person_primary_role(person):
    # determine if person is ever Night in month -> put at bottom; else keep order as in members
    for w in weeks:
        roles = week_role_assignment[w]
        if roles["Night"] == person:
            return "Night"
    # else return "Other" to keep earlier positions
    return "Other"

# create ordering: non-night in original members order, then night people in original members order
non_night = [m for m in members if person_primary_role(m) != "Night"]
night_people = [m for m in members if person_primary_role(m) == "Night"]
final_order = non_night + night_people

df_display = df_display.set_index("Name").loc[final_order].reset_index()

# ---------------- Display ----------------
st.subheader(f"Rota for {date(year, month, 1).strftime('%B %Y')} — mode: {'Mon–Sun' if seven_day_mode else 'Mon–Fri'}")
st.markdown("Rows ordered with Night shift people at the bottom. `UNAVAILABLE` means the person had full-week leave or no suitable assignment.")

# show a styled dataframe
st.dataframe(df_display, 1000, 600)

# show the week-role mapping for transparency
st.markdown("### Week role mapping (chosen persons per week)")
wk_rows = []
for i, w in enumerate(weeks):
    r = {"Week": f"Week {i+1} (start {w.strftime('%Y-%m-%d')})"}
    r.update(week_role_assignment[w])
    wk_rows.append(r)
st.table(pd.DataFrame(wk_rows))

# ---------------- Downloads ----------------
csv = df_display.to_csv(index=False).encode("utf-8")
st.download_button("Download CSV", csv, file_name=f"rota_{year}_{month:02d}.csv", mime="text/csv")

if download_excel:
    try:
        import io
        with io.BytesIO() as output:
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_display.to_excel(writer, index=False, sheet_name="Rota")
                # second sheet: week mapping
                pd.DataFrame(wk_rows).to_excel(writer, index=False, sheet_name="WeekMapping")
            data = output.getvalue()
        st.download_button("Download Excel", data, file_name=f"rota_{year}_{month:02d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error("Excel export failed (missing openpyxl?). Install openpyxl to enable Excel export.")

