import os

import streamlit as st
import openpyxl
from icalendar import Calendar, Event
import datetime
import hashlib
import parsedatetime as pdt
import re
import pandas as pd
from xml.sax.saxutils import escape
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, PageBreak


def format_selection(university):
    uni_format = None

    if university == "University of Newcastle Y2 2026":
        uni_format = {"preset": "University of Newcastle Y2 2026",
                      "uni": "UON",
                      "day": 1,
                      "date": 2,
                      "time": 3,
                      "campus": 5,
                      "group": 6,
                      "venue": 7,
                      "attendance": 8,
                      "type": 9,
                      "domain": 10,
                      "session": 11,
                      "staff": 12,
                      "updates": 13,
                      "wp_from": "sessions",
                      "default_zoom": "https://uonewcastle.zoom.us/j/89060830849",
                      "fields": ["pbl"]
                    }


    if university == "University of New England Y2 2026":
        uni_format = {
                      "preset": "University of New England Y2 2026",
                      "uni": "UNE",
                      "day": 0,
                      "date": 1,
                      "time": 2,
                      "campus": 15,
                      "group": 4,
                      "venue": 5,
                      "attendance": 6,
                      "type": 7,
                      "domain": 8,
                      "session": 9,
                      "staff": 10,
                      "updates": 11,
                      "wp_from": "headings",
                      "default_zoom": "https://uonewcastle.zoom.us/j/89060830849",
                      "fields": ["pbl", "clin", "comm"]
                      }
    return uni_format

def main_ui():
    st.set_page_config(page_title="JMPCalendarConverter")

    st.title("JMPCalendarConverter")
    st.text("Works well enough for now, I'm still ironing out bugs. Lmk if you notice anything @photogrudesh.")
    with st.status("What's new"):
        st.write("Added spreadsheet preview to see detected events")
        st.write("Added spreadsheet downloads")
        st.write("Removed custom mode cause lowkey who is even bothered.")
        st.write("Add default zoom link to events with failed hyperlink extraction.")
        st.write("Tidy up redundant code")
        st.write("This message 😱")
        st.write("Toast notification on completion")
        st.write("Balloon 🎈")

    st.divider()

    if not st.toggle("Use your own file"):
        university = st.selectbox("Preset calendar - Year 2 only rn - DM @photogrudesh if this is outdated.", ["University of Newcastle Y2 2026", "University of New England Y2 2026"])
    else:
        university = "Use your own file"

    file = None

    files = os.listdir()
    uoncals = []
    unecals = []

    for f in files:
        if f.endswith(".xlsx"):
            if f.startswith("CALCCCS"):
                uoncals.append(f)
            elif f.startswith("UNEARM"):
                unecals.append(f)

    file_select, file_dl = st.columns([5, 2])

    if university == "Use your own file":
        file = st.file_uploader("Upload your own calendar file", type=None, accept_multiple_files=False)
        uni_format_select = st.radio("Format",
                                     ["University of Newcastle Y2 2026", "University of New England Y2 2026"],
                                     horizontal=True)

        uni_format = format_selection(uni_format_select)
    else:
        uni_format = format_selection(university)
        with file_select:
            if uni_format:
                if uni_format["uni"] == "UON":
                    file = st.selectbox("uon_file", uoncals, label_visibility="collapsed")
                elif uni_format["uni"] == "UNE":
                    file = st.selectbox("une_file", unecals, label_visibility="collapsed")
        with file_dl:
            with open(file, "rb") as f:
                st.download_button(
                    label="Download spreadsheet",
                    data=f,
                    file_name=file,
                    width="stretch",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


    valid_custom = True


    date_start, date_end = None, None

    valid_selection = False

    if file and valid_custom:
        # process file and determine which date ranges are valid
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        st.divider()
        st.subheader("Options")

        option = st.radio("Convert", ["All events", "Events after date", "Custom dates"], horizontal=True)

        campus = None
        pbl = None
        clin = None
        comm = None

        column1, column2 = st.columns(2)

        if option == "All events":
            date_start = datetime.date(1970, 1, 1)
            date_end = datetime.date(2099, 1, 1)
        elif option == "Events after date":
            with column1:
                date_start = st.date_input("Date to import from (inclusive)", value="today", min_value=None,
                                           max_value=None,
                                           format="DD/MM/YYYY", key="dsi")
                date_end = datetime.date(2099, 1, 1)

        elif option == "Custom dates":
            with column1:
                date_start = st.date_input("Date to import from (inclusive)", value="today", min_value=None,
                                           max_value=None,
                                           format="DD/MM/YYYY", key="dsi")
            with column2:
                date_end = st.date_input("Last date to import (inclusive)", value="today", min_value=None,
                                         max_value=None,
                                         format="DD/MM/YYYY", key="dei")


        if uni_format["uni"] == "UON":
            campus = st.selectbox("Campus", ["Callaghan", "Central Coast"])

        colpbl, colclin, colcomm = st.columns(3)

        if not st.toggle("Manual entry"):
            with colpbl:
                if "pbl" in uni_format["fields"]:
                    pbl = st.selectbox("PBL group", ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"), index=None, placeholder="Select PBL group")
                else:
                    pbl = "z"
            with colclin:
                if "clin" in uni_format["fields"]:
                    clin = st.selectbox("Clin group", ("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"), index=None, placeholder="Select clin group")
                else:
                    clin = "300"
            with colcomm:
                if "comm" in uni_format["fields"]:
                    comm = st.selectbox("Comm group", ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"), index=None, placeholder="Select comm group")
                else:
                    comm = "z"
        else:
            with colpbl:
                if "pbl" in uni_format["fields"]:
                    pbl = st.text_input("PBL group", placeholder="Select PBL group")
                else:
                    pbl = "z"
            with colclin:
                if "clin" in uni_format["fields"]:
                    clin = st.text_input("Clin group", placeholder="Select clin group")
                else:
                    clin = "300"
            with colcomm:
                if "comm" in uni_format["fields"]:
                    comm = st.text_input("Comm group", placeholder="Select comm group")
                else:
                    comm = "z"

        previous = None
        pdf_theme = st.radio("PDF style", ["Light", "Dark"], horizontal=True)

        changes_only = st.toggle("Only export changes since my last import")

        if changes_only:
            prev_source = st.radio("Compare against", ["Spreadsheet I used last time", "Calendar file (.ics) I downloaded last time"], horizontal=True)
            if prev_source == "Spreadsheet I used last time":
                cals = uoncals if uni_format["uni"] == "UON" else unecals
                older = [c for c in cals if c != file]
                prev_file = st.selectbox("Previous spreadsheet", older + ["Upload a spreadsheet"])
                if prev_file == "Upload a spreadsheet":
                    prev_file = st.file_uploader("Upload the spreadsheet you used last time", type="xlsx", key="prev_xlsx")
                if prev_file:
                    previous = ("xlsx", prev_file)
            else:
                prev_ics = st.file_uploader("Upload the full calendar (.ics) you downloaded last time", type="ics", key="prev_ics")
                if prev_ics:
                    previous = ("ics", prev_ics)

        if option == "All events":
            st.text(f"Importing all available events from {file}")

        if clin and comm and pbl and (previous or not changes_only):
            valid_selection = True

        go = False


        if valid_selection and pbl:
            pre_col, go_col = st.columns([1, 4])
            with pre_col:
                preview = st.button("Preview", use_container_width=True)
            with go_col:
                go = st.button("Start converting", use_container_width=True)
            st.text("Always check the output below for errors and double check calendar events with the spreadsheet on canvas. Scroll down to download the file when it's ready.")
        elif clin and comm and pbl:
            st.warning("Pick the spreadsheet or calendar file you used last time.")
        else:
            st.warning(f"Select your timetable.")

        if valid_selection and pbl:
            if preview:
                rows_kept = process_xlsx(pbl.upper(), clin, comm, uni_format, campus, ws, True, date_start, date_end)
                xlsx_preview(file, rows_kept)


        if go and valid_selection:
            saved, rows_kept = process_xlsx(pbl.upper(), clin, comm, uni_format, campus, ws, False, date_start, date_end)
            xlsx_preview(file, rows_kept)

            prev_events = None
            if previous:
                prev_events = load_previous(previous, pbl.upper(), clin, comm, uni_format, campus, date_start, date_end)

            groups = []
            if "pbl" in uni_format["fields"]:
                groups.append(f"PBL {pbl.upper()}")
            if "clin" in uni_format["fields"]:
                groups.append(f"Clin {clin}")
            if "comm" in uni_format["fields"]:
                groups.append(f"Comm {comm.upper()}")
            if campus:
                groups.append(campus)

            converted = generate_cal(saved, date_start, date_end, uni_format, prev_events, "  ·  ".join(groups), pdf_theme)

            if os.path.exists("calendar.ics"):
                f = open("calendar.ics", "r")

                if prev_events is None:
                    st.download_button("Download ics file", data=f, file_name="Calendar.ics", use_container_width=True)

                    st.text(
                        f"Import this file to your calendar app (google calendar works idk about the rest)\nAlways double check to see if events have been imported correctly. {converted} events were converted by the converter. If google calendar doesn't import the same number of events, check the spreadsheet. DM me @photogrudesh on Instagram if there are any issues.")
                else:
                    if converted:
                        st.download_button("Download changes only", data=f, file_name="Calendar changes.ics", use_container_width=True)
                    full = open("calendar_full.ics", "r")
                    st.download_button("Download full calendar (keep this to compare against next time)", data=full, file_name="Calendar.ics", use_container_width=True)

                    st.text(
                        f"Import the changes file to your calendar app. {converted} new or changed events were exported. If a changed event shows up twice, delete the old one. Removed events have to be deleted from your calendar yourself. DM me @photogrudesh on Instagram if there are any issues.")

                if os.path.exists("timetable.pdf"):
                    pdf = open("timetable.pdf", "rb")
                    st.download_button("Download pdf timetable", data=pdf, file_name="Timetable.pdf", mime="application/pdf", use_container_width=True)
                st.toast("Calendar converted, scroll down to download", icon="🥳")
                st.balloons()



# rows are read 20 columns wide, so the spreadsheet's week label gets tacked on after them
WEEK_LABEL = 20


def spreadsheet_weeks(rows, uni_format):
    labels = [None] * len(rows)

    if uni_format["wp_from"] == "sessions":
        # UON has no WP headings, but sessions are named "Open Working Problem 45" etc.
        # each WP runs from the first day it comes up until the next one starts
        starts = {}
        wp_weeks = set()
        for row in rows:
            date = row[uni_format["date"]]
            match = re.search(r"Working Prob\w*\s*(\d+)|WP\s*(\d+)", str(row[uni_format["session"]] or ""), re.I)
            if isinstance(date, datetime.datetime) and match:
                number = int(match.group(1) or match.group(2))
                starts[number] = min(starts.get(number, date), date)
                wp_weeks.add(date.date() - datetime.timedelta(days=date.weekday()))

        ordered = sorted(starts.items(), key=lambda x: x[1])
        for n, row in enumerate(rows):
            date = row[uni_format["date"]]
            if not isinstance(date, datetime.datetime):
                continue
            # weeks with no WP at all (GP placement, vacation) get a plain "Week of" heading
            if date.date() - datetime.timedelta(days=date.weekday()) not in wp_weeks:
                continue
            current = [number for number, start in ordered if start <= date]
            if current:
                labels[n] = f"WP {current[-1]}"
            elif ordered:
                # the first days of semester come before the first WP opens
                labels[n] = f"WP {ordered[0][0]}"
    else:
        # UNE splits the sheet into sections with a heading row, e.g. "WP 45: Hormones and the Lifecycle"
        current = None
        for n, row in enumerate(rows):
            if row[0] and not row[uni_format["date"]]:
                current = str(row[0]).strip()
            labels[n] = current

    return labels


def process_xlsx(pbl, clin, comm, uni_format, campus, ws, preview, date_start, date_end):
    events = []
    saved = []
    rows_kept = []
    preview_error = False

    row_num = 1

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=20, values_only=True):
        events.append(row)

    week_labels = spreadsheet_weeks(events, uni_format)

    for i in events:
        i = list(i)
        i.append(week_labels[row_num - 1])
        keep = False
        group = str(i[uni_format["group"]]).strip()
        complex_grp = group.split(" ")[-1]

        if "all" in group.lower():
            keep = True
        elif group.startswith("PBL"):
            if bool(re.search(rf'(?<!\w){re.escape(pbl)}(?!\w)', group)):
                keep = True
            if "-" in complex_grp:
                start_char = complex_grp.split("-")[0].strip()
                end_char = complex_grp.split("-")[1].strip()
                for x in [chr(c) for c in range(ord(start_char), ord(end_char) + 1)]:
                    if pbl == str(x):
                        keep = True
            elif "," in group:
                pbls = group.split(",")
                for f in pbls:
                    if f.strip() == pbl:
                        keep = True
        elif len(group) == 1:
            if group == pbl:
                keep = True

        elif group.startswith("Clin"):
            if bool(re.search(rf'(?<!\w){re.escape(clin)}(?!\w)', group)):
                keep = True
            if "-" in complex_grp:
                start, end = complex_grp.split("-")
                start_num = ""
                end_num = ""
                clin_groups = []
                for n in start:
                    if n.isdigit():
                        start_num += n
                start_num = int(start_num)
                for n in end:
                    if n.isdigit():
                        end_num += n
                end_num = int(end_num)
                for n in range(start_num, end_num + 1):
                    clin_groups.append(str(n))
                if clin in clin_groups:
                    keep = True

            elif "," in group:
                pbls = group.split(",")
                for f in pbls:
                    if f.strip() == pbl:
                        keep = True

        elif group.startswith("Comms"):
            if bool(re.search(rf'(?<!\w){re.escape(comm)}(?!\w)', group)):
                keep = True
            if "-" in complex_grp:
                start_char = complex_grp.split("-")[0].strip()
                end_char = complex_grp.split("-")[1].strip()
                for x in [chr(c) for c in range(ord(start_char), ord(end_char) + 1)]:
                    if pbl == str(x):
                        keep = True
            elif "," in group:
                comms = group.split(",")
                for f in comms:
                    if f.strip() == comm:
                        keep = True
        else:
            if "-" in complex_grp:
                start_char = complex_grp.split("-")[0].strip()
                end_char = complex_grp.split("-")[1].strip()
                for x in [chr(c) for c in range(ord(start_char), ord(end_char) + 1)]:
                    if pbl == str(x):
                        keep = True
            elif "," in group:
                pbls = group.split(",")
                for f in pbls:
                    if f.strip() == pbl:
                        keep = True


        if "zoom" in str(i[uni_format["venue"]]).lower():
            # openpyxl columns are 1-indexed, row values are 0-indexed
            zoom_col = uni_format["venue"] + 1
            try:
                target = ws.cell(row=row_num, column=zoom_col).hyperlink.target
                # some cells link to a file on the uni network drive instead of zoom
                if not str(target).startswith("http"):
                    raise AttributeError
                i[uni_format["venue"]] = target
            except AttributeError:
                i[uni_format["venue"]] = str(f"{uni_format["default_zoom"]}\n" + i[uni_format["venue"]]) + f": Failed hyperlink extraction. Default JMP zoom added instead. Check spreadsheet/canvas."

        if uni_format["preset"] == "University of Newcastle Y2 2026":
            campus_code = None
            if campus == "Callaghan":
                campus_code = "CAL"
            elif campus == "Central Coast":
                campus_code = "CC"
            if campus_code not in str(i[uni_format["campus"]]):
                keep = False

        elif not uni_format["preset"]:
            campus_code = uni_format["campus_code"]
            if campus_code not in str(i[uni_format["campus"]]):
                keep = False

        if keep:
            if not preview:
                saved.append(i)

            try:
                if date_start - datetime.timedelta(days=1) < i[uni_format["date"]].date() < date_end + datetime.timedelta(days=1):
                    rows_kept.append(row_num - 2)
            except AttributeError:
                preview_error = True

        row_num += 1

    if preview_error:
        st.warning("Preview may not be accurate. Download the spreadsheet to double check events.")

    if preview:
        return rows_kept
    else:
        return saved, rows_kept


def highlight_selected_rows(df, rows):
    rows = set(rows)
    return df.style.apply(
        lambda row: ["background-color: #fff3b0; color: #000000"] * len(row) if row.name in rows else [""] * len(row),
        axis=1,
    )


def xlsx_preview(file, rows):
    df = pd.read_excel(file)
    # st.dataframe(df, use_container_width=True)
    st.dataframe(highlight_selected_rows(df.iloc[rows[0]:], rows), use_container_width=True)
    if len(rows) == 1:
        st.success(f"1 relevant event will be converted.")

    else:
        st.success(f"{len(rows)} relevant events will be converted.")


def generate_cal(events, date_start, date_end, uni_format, previous=None, subtitle="", pdf_theme="Light"):
    new_events = build_events(events, date_start, date_end, uni_format)

    if previous is None:
        show_events("Importing events...", new_events)
        with open("calendar.ics", "wb") as f:
            f.write(to_ics(new_events))
        generate_pdf(new_events, uni_format, subtitle, theme=pdf_theme)
        return len(new_events)

    added, changed, removed = diff_events(new_events, previous)

    for old, new in changed:
        # tells calendar apps this is a newer version of an event they already have
        new.add("sequence", int(old.get("sequence", 0)) + 1)

    if not (added or changed or removed):
        st.info("No changes since last time.")

    show_events(f"New events ({len(added)})", added)

    with st.status(f"Changed events ({len(changed)})", expanded=True):
        for old, new in changed:
            with st.status(event_label(new)):
                st.write("**Now:**")
                st.write(str(new.get("description", "")))
                st.write(f"**Before:** {event_label(old)}")
                st.write(str(old.get("description", "")))

    show_events(f"Removed events ({len(removed)}) - delete these from your calendar yourself", removed)

    exported = added + [new for old, new in changed]
    # calendar.ics gets just the changes, calendar_full.ics is kept to compare against next time
    with open("calendar.ics", "wb") as f:
        f.write(to_ics(exported))
    with open("calendar_full.ics", "wb") as f:
        f.write(to_ics(new_events))
    generate_pdf(new_events, uni_format, subtitle, (added, changed, removed), pdf_theme)
    return len(exported)


def load_previous(previous, pbl, clin, comm, uni_format, campus, date_start, date_end):
    kind, source = previous

    if kind == "xlsx":
        ws = openpyxl.load_workbook(source).active
        saved, rows_kept = process_xlsx(pbl, clin, comm, uni_format, campus, ws, False, date_start, date_end)
        return build_events(saved, date_start, date_end, uni_format, show_errors=False)

    cal = Calendar.from_ical(source.getvalue())
    events = []
    for e in cal.walk("VEVENT"):
        start = e.decoded("dtstart")
        day = start.date() if isinstance(start, datetime.datetime) else start
        # only compare against events in the date range being exported
        if date_start <= day <= date_end:
            events.append(e)
    return events


def fingerprint(e):
    dtend = e.decoded("dtend") if "dtend" in e else None
    return str(e.get("summary", "")), e.decoded("dtstart"), dtend, str(e.get("description", ""))


def diff_events(new, old):
    # ics files from before UIDs were added can only be matched on title + start time,
    # so a changed event from one of those shows up as removed + new
    if not old or all("uid" in e for e in old):
        key = lambda e: str(e["uid"])
    else:
        key = lambda e: (str(e.get("summary", "")), e.decoded("dtstart"))

    old_by_key = {key(e): e for e in old}
    new_keys = set()
    added = []
    changed = []

    for e in new:
        k = key(e)
        new_keys.add(k)
        if k not in old_by_key:
            added.append(e)
        elif fingerprint(e) != fingerprint(old_by_key[k]):
            changed.append((old_by_key[k], e))

    removed = [e for k, e in old_by_key.items() if k not in new_keys]
    return added, changed, removed


def event_label(e):
    summary = str(e.get("summary", "")).replace('\n', ' ')
    start = e.decoded("dtstart")
    if isinstance(start, datetime.datetime):
        return f"{start}: " + summary
    return summary


def show_events(title, events):
    with st.status(title, expanded=True):
        for e in events:
            with st.status(event_label(e)):
                st.write(str(e.get("description", "")))


def to_ics(events):
    cal = Calendar()
    cal['version'] = '2.0'
    cal.add("prodid", "-//photogrudesh//JMPCalendarConverter//EN")
    cal.add("summary", "JMP schedule")
    for e in events:
        cal.add_component(e)
    return cal.to_ical()


def build_events(events, date_start, date_end, uni_format, show_errors=True):
    built = []
    used_uids = set()
    dtstamp = datetime.datetime.now(datetime.timezone.utc)

    for i in events:
        try:
            if isinstance(i[uni_format["date"]], str):
                # dates typed as text, e.g. "3/8/2026" or "18/08'/2026"
                date_text = re.sub(r"[^\d/]", "", i[uni_format["date"]].replace("-", "/"))
                event_date = datetime.datetime.strptime(date_text, "%d/%m/%Y")
            else:
                event_date = i[uni_format["date"]]

            if date_start - datetime.timedelta(days=1) < event_date.date() < date_end + datetime.timedelta(days=1):
                event = Event()
                no_time = False

                try:
                    start, end = convert_datetime(event_date, i[uni_format["time"]])
                    event.add('dtstart', start)
                    event.add('dtend', end)
                except TypeError:
                    no_time = True

                if i[uni_format["attendance"]] is None:
                    attendance = "N/A"
                else:
                    attendance = i[uni_format["attendance"]]

                desc = f"{i[uni_format["type"]]}: {i[uni_format["domain"]]}\n{i[uni_format["venue"]]}\nStudents: {i[uni_format["group"]]}\nAttendance: {attendance}\nStaff: {i[uni_format["staff"]]}\nUpdates: {i[uni_format["updates"]]}"

                if no_time:
                    # date (not datetime) makes it a proper all-day event
                    event.add('dtstart', event_date.date())
                    event.add('dtend', event_date.date() + datetime.timedelta(days=1))

                summary = str(i[uni_format["session"]] or "").strip()
                venue = str(i[uni_format["venue"]] or "").strip()

                if "MAND" in str(attendance):
                    summary = "M - " + summary
                if "zoom" not in venue.lower():
                    if venue:
                        summary = summary + f" - {venue}"
                else:
                    summary = summary + f" - ZOOM"

                event.add('summary', summary)
                event.add("description", desc)
                event.add('uid', make_uid(uni_format, event_date, i, used_uids))
                event.add('dtstamp', dtstamp)
                # plain details for the pdf, so it doesn't have to pick apart the summary
                event.info = {
                    "title": str(i[uni_format["session"]] or "").replace("\n", " ").strip(),
                    "venue": "Zoom" if "zoom" in venue.lower() else venue.replace("\n", " "),
                    "type": str(i[uni_format["type"]] or "").replace("\n", " ").strip(),
                    "mandatory": "MAND" in str(attendance),
                    "week": i[WEEK_LABEL] if len(i) > WEEK_LABEL else None,
                }
                built.append(event)
        except (AttributeError, ValueError) as e:
            print(e)
            if show_errors:
                st.error(f"Something went wrong. Check the spreadsheet for information on {i}\n{e}")
                st.toast(f"Something went wrong. Add {i} manually.", duration="infinite")

    return built


# every text colour is at least 4.5:1 against every background it sits on in dark mode
PDF_THEMES = {
    "Light": {
        "bg": "#ffffff", "ink": "#1d1d1f", "grey": "#6e6e73", "light_grey": "#86868b",
        "hairline": "#d2d2d7", "card": "#f5f5f7",
        "sage": "#5b7b62", "sage_tint": "#eef3ec", "orange": "#b64400", "orange_tint": "#fbf1e8", "red": "#d70015",
    },
    "Dark": {
        "bg": "#161617", "ink": "#f5f5f7", "grey": "#aeaeb2", "light_grey": "#98989d",
        "hairline": "#3a3a3c", "card": "#232326",
        "sage": "#9cc0a2", "sage_tint": "#1f2b22", "orange": "#f5a35c", "orange_tint": "#33261b", "red": "#ff6961",
    },
}


def pdf_styles(t):
    c = lambda key: colors.HexColor(t[key])
    return {
        "eyebrow": ParagraphStyle("eyebrow", fontName="Helvetica-Bold", fontSize=10, leading=14, textColor=c("sage")),
        "tag": ParagraphStyle("tag", fontName="Helvetica-Bold", fontSize=7.5, leading=10.5),
        "headline": ParagraphStyle("headline", fontName="Helvetica-Bold", fontSize=40, leading=43, textColor=c("ink")),
        "sub": ParagraphStyle("sub", fontName="Helvetica", fontSize=13, leading=19, textColor=c("grey")),
        "stat": ParagraphStyle("stat", fontName="Helvetica-Bold", fontSize=28, leading=32, textColor=c("ink")),
        "stat_label": ParagraphStyle("stat_label", fontName="Helvetica", fontSize=9, leading=12, textColor=c("grey")),
        "section": ParagraphStyle("section", fontName="Helvetica-Bold", fontSize=24, leading=28, textColor=c("ink")),
        "section_sub": ParagraphStyle("section_sub", fontName="Helvetica", fontSize=11, leading=15, textColor=c("grey")),
        "day_name": ParagraphStyle("day_name", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=c("light_grey")),
        "day_num": ParagraphStyle("day_num", fontName="Helvetica-Bold", fontSize=20, leading=23, textColor=c("ink")),
        "time": ParagraphStyle("time", fontName="Helvetica", fontSize=9, leading=12, textColor=c("grey")),
        "title": ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=10.5, leading=13.5, textColor=c("ink")),
        "meta": ParagraphStyle("meta", fontName="Helvetica", fontSize=8.5, leading=11.5, textColor=c("grey")),
        "card_label": ParagraphStyle("card_label", fontName="Helvetica-Bold", fontSize=10, leading=13),
        "card_item": ParagraphStyle("card_item", fontName="Helvetica", fontSize=9, leading=12.5, textColor=c("ink")),
    }


def pdf_time(t):
    return f"{t:%I:%M %p}".lstrip("0").lower()


def pdf_day(d):
    return f"{d:%a} {d.day} {d:%B}"


def pdf_when(e):
    start = e.decoded("dtstart")
    if isinstance(start, datetime.datetime):
        return f"{pdf_day(start)}, {pdf_time(start)}"
    return f"{pdf_day(start)}, all day"


def pdf_title(e):
    if hasattr(e, "info"):
        return e.info["title"]
    return str(e.get("summary", "")).replace("\n", " ")


def pdf_page(t):
    def draw(canvas, doc):
        canvas.saveState()
        # page background first, so dark mode isn't just dark text boxes on white paper
        canvas.setFillColor(colors.HexColor(t["bg"]))
        canvas.rect(0, 0, A4[0], A4[1], stroke=0, fill=1)
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.HexColor(t["light_grey"]))
        canvas.drawString(doc.leftMargin + 6, 12 * mm,
                          f"Generated {datetime.date.today().day} {datetime.date.today():%B %Y} by JMPCalendarConverter. Always double check against Canvas.")
        canvas.drawRightString(A4[0] - doc.rightMargin - 6, 12 * mm, str(doc.page))
        canvas.restoreState()
    return draw


def pdf_changes(changes, width, t, s):
    added, changed, removed = changes
    flow = [Paragraph("What's changed.", s["section"]), Spacer(1, 3 * mm),
            Paragraph("Since the last version you imported. New and changed events are highlighted in the timetable too.", s["section_sub"]),
            Spacer(1, 7 * mm)]

    if not (added or changed or removed):
        flow.append(Paragraph("Nothing. You're all up to date.", s["sub"]))

    groups = [
        ("New", t["sage"], [f"{escape(pdf_when(e))}  ·  <b>{escape(pdf_title(e))}</b>" for e in added]),
        ("Changed", t["orange"], [f"{escape(pdf_when(new))}  ·  <b>{escape(pdf_title(new))}</b><br/>"
                                  f"<font color='{t['grey']}'>Was {escape(pdf_when(old))}  ·  {escape(pdf_title(old))}</font>"
                                  for old, new in changed]),
        ("Removed", t["red"], [f"<strike>{escape(pdf_when(e))}  ·  {escape(pdf_title(e))}</strike>" for e in removed]),
    ]

    for label, colour, items in groups:
        if not items:
            continue
        label_style = ParagraphStyle(f"label_{label}", parent=s["card_label"], textColor=colors.HexColor(colour))
        rows = [[Paragraph(f"{label}  <font color='{t['light_grey']}'>{len(items)}</font>", label_style)]]
        rows += [[Paragraph(item, s["card_item"])] for item in items]
        card = Table(rows, colWidths=[width])
        card.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(t["card"])),
            ("ROUNDEDCORNERS", [10, 10, 10, 10]),
            ("LEFTPADDING", (0, 0), (-1, -1), 6 * mm),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6 * mm),
            ("TOPPADDING", (0, 0), (-1, -1), 2.2 * mm),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2 * mm),
            ("TOPPADDING", (0, 0), (-1, 0), 5 * mm),
            ("BOTTOMPADDING", (0, -1), (-1, -1), 5 * mm),
        ]))
        flow += [card, Spacer(1, 5 * mm)]

    flow.append(PageBreak())
    return flow


def pdf_day_block(day, day_events, width, highlights, t, s):
    time_w = 24 * mm
    day_w = 20 * mm
    pad = 2.5 * mm

    rows = []
    for e in day_events:
        start = e.decoded("dtstart")
        if isinstance(start, datetime.datetime):
            end = e.decoded("dtend")
            when = f"{pdf_time(start)}<br/><font color='{t['light_grey']}'>{pdf_time(end)}</font>"
        else:
            when = "All day"

        info = e.info
        meta = "  ·  ".join(escape(x) for x in [info["type"], info["venue"]] if x)
        if info["mandatory"]:
            meta = f"<font color='{t['sage']}'><b>Mandatory</b></font>" + (f"  ·  {meta}" if meta else "")

        content = []
        highlight = highlights.get(str(e.get("uid")))
        tint = None
        if highlight:
            status, old = highlight
            if status == "new":
                tint = t["sage_tint"]
                content.append(Paragraph(f"<font color='{t['sage']}'>NEW</font>", s["tag"]))
            else:
                tint = t["orange_tint"]
                content.append(Paragraph(f"<font color='{t['orange']}'>CHANGED</font>"
                                         f"<font name='Helvetica' color='{t['grey']}'>  ·  was {escape(pdf_when(old))}</font>", s["tag"]))
        content.append(Paragraph(escape(info["title"]), s["title"]))
        if meta:
            content.append(Paragraph(meta, s["meta"]))

        # every event gets the same padding so highlighted and plain ones line up
        row = Table([[Paragraph(when, s["time"]), content]], colWidths=[time_w, width - day_w - time_w - 2 * pad])
        row_style = [
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]
        row.setStyle(TableStyle(row_style))
        card = Table([[row]], colWidths=[width - day_w])
        card_style = [
            ("LEFTPADDING", (0, 0), (-1, -1), pad),
            ("RIGHTPADDING", (0, 0), (-1, -1), pad),
            ("TOPPADDING", (0, 0), (-1, -1), pad),
            ("BOTTOMPADDING", (0, 0), (-1, -1), pad),
        ]
        if tint:
            card_style += [("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(tint)), ("ROUNDEDCORNERS", [8, 8, 8, 8])]
        card.setStyle(TableStyle(card_style))
        rows.append([card])

    events_table = Table(rows, colWidths=[width - day_w])
    events_table.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1 * mm),
    ]))

    day_cell = [Paragraph(f"{day:%a}".upper(), s["day_name"]), Paragraph(str(day.day), s["day_num"])]
    block = Table([[day_cell, events_table]], colWidths=[day_w, width - day_w])
    block.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        # events carry their own padding, so the day number gets it here to stay level with them
        ("TOPPADDING", (0, 0), (0, 0), 1.5 * mm + pad),
        ("TOPPADDING", (1, 0), (1, 0), 1.5 * mm),
        # a day with one event is about as tall as the date, so give the date room before the next divider
        ("BOTTOMPADDING", (0, 0), (0, 0), 3.5 * mm),
        ("BOTTOMPADDING", (1, 0), (1, 0), 0.5 * mm),
        ("LINEABOVE", (0, 0), (-1, 0), 0.5, colors.HexColor(t["hairline"])),
    ]))
    return block


def generate_pdf(events, uni_format, subtitle, changes=None, theme="Light"):
    t = PDF_THEMES[theme]
    s = pdf_styles(t)
    doc = SimpleDocTemplate("timetable.pdf", pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm,
                            topMargin=22 * mm, bottomMargin=22 * mm, title="JMP Timetable", author="JMPCalendarConverter")
    # the page frame has 6pt of padding each side, tables need to fit inside it to line up with the text
    width = doc.width - 12

    def event_day(e):
        start = e.decoded("dtstart")
        return start.date() if isinstance(start, datetime.datetime) else start

    def event_sort(e):
        start = e.decoded("dtstart")
        if not isinstance(start, datetime.datetime):
            start = datetime.datetime.combine(start, datetime.time.min)
        return start

    events = sorted(events, key=event_sort)
    days = sorted({event_day(e) for e in events})

    flow = [Paragraph(escape(uni_format["preset"].replace(" Y2 ", " · Year 2 · ")), s["eyebrow"]),
            Spacer(1, 4 * mm),
            Paragraph("Your semester.<br/>All in one place.", s["headline"]),
            Spacer(1, 6 * mm)]
    if subtitle:
        flow.append(Paragraph(escape(subtitle), s["sub"]))
    if days:
        flow.append(Paragraph(f"{days[0].day} {days[0]:%B} – {days[-1].day} {days[-1]:%B %Y}", s["sub"]))
    flow.append(Spacer(1, 12 * mm))

    weeks = sorted({d - datetime.timedelta(days=d.weekday()) for d in days})
    stats = [(len(events), "sessions"),
             (sum(1 for e in events if e.info["mandatory"]), "mandatory"),
             (len(weeks), "weeks")]
    stats_table = Table([[Paragraph(str(n), s["stat"]) for n, label in stats],
                         [Paragraph(label, s["stat_label"]) for n, label in stats]],
                        colWidths=[width / 3] * 3)
    stats_table.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (-1, 0), 0.5, colors.HexColor(t["hairline"])),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, 0), 5 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1 * mm),
    ]))
    flow += [stats_table, Spacer(1, 14 * mm)]

    if changes is not None:
        flow += pdf_changes(changes, width, t, s)

    # uid -> ("new", None) or ("changed", old event), to mark them in the timetable
    highlights = {}
    if changes is not None:
        added, changed, removed = changes
        for e in added:
            highlights[str(e.get("uid"))] = ("new", None)
        for old, new in changed:
            highlights[str(new.get("uid"))] = ("changed", old)

    # sections follow the spreadsheet's working problems in date order. a WP that gets
    # interrupted (e.g. by GP placement) picks up again afterwards as "continued"
    sections = []
    seen = set()
    for e in events:
        label = e.info["week"]
        if not label:
            monday = event_day(e) - datetime.timedelta(days=event_day(e).weekday())
            label = f"Week of {monday.day} {monday:%B}"
        if not sections or sections[-1][0] != label:
            sections.append((label, [], label in seen))
            seen.add(label)
        sections[-1][1].append(e)

    for label, section_events, continued in sections:
        week_days = sorted({event_day(e) for e in section_events})
        first, last = week_days[0], week_days[-1]
        if first == last:
            when = f"{first.day} {first:%B}"
        elif first.month == last.month:
            when = f"{first.day} – {last.day} {last:%B}"
        else:
            when = f"{first.day} {first:%B} – {last.day} {last:%B}"

        # "WP 45: Hormones and the Lifecycle" -> "WP 45" with the topic underneath
        title, sep, topic = re.sub(r"^WP\s*(\d+)", r"WP \1", label).partition(":")
        sub = "  ·  ".join(x for x in ["continued" if continued else "", topic.strip(), when] if x)

        heading = [Paragraph(escape(title.strip()), s["section"]), Spacer(1, 1.5 * mm),
                   Paragraph(escape(sub), s["section_sub"]), Spacer(1, 4 * mm)]
        for i, day in enumerate(week_days):
            block = pdf_day_block(day, [e for e in section_events if event_day(e) == day], width, highlights, t, s)
            # keep the heading with its first day so it never sits alone at the bottom of a page
            flow.append(KeepTogether(heading + [block]) if i == 0 else KeepTogether([block]))
        flow.append(Spacer(1, 10 * mm))

    doc.build(flow, onFirstPage=pdf_page(t), onLaterPages=pdf_page(t))


def make_uid(uni_format, event_date, i, used_uids):
    # stable across conversions so re-importing updates events instead of duplicating them
    key = "|".join([uni_format["uni"], event_date.date().isoformat(),
                    str(i[uni_format["session"]]), str(i[uni_format["group"]])])
    uid = hashlib.sha1(key.encode()).hexdigest()
    n = 1
    base = uid
    while uid in used_uids:
        n += 1
        uid = f"{base}-{n}"
    used_uids.add(uid)
    return f"{uid}@jmpcalendarconverter"


def convert_datetime(date, time):
    time = time.replace(".", ":").replace("md", "pm")

    times = time.split(" - ")

    try:
        cal = pdt.Calendar()
        time_struct, parse_status = cal.parse(times[0])
        start_time = datetime.datetime(*time_struct[:6]).time()
        cal = pdt.Calendar()
        time_struct, parse_status = cal.parse(times[1])
        end_time = datetime.datetime(*time_struct[:6]).time()

        start = datetime.datetime.combine(date, start_time)
        end = datetime.datetime.combine(date, end_time)

    except IndexError:
        return date
    return start, end


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main_ui()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
