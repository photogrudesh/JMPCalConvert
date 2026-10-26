import os

import streamlit as st
import openpyxl
from icalendar import Calendar, Event
import datetime
import hashlib
import parsedatetime as pdt
import re
import pandas as pd


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

            converted = generate_cal(saved, date_start, date_end, uni_format, prev_events)

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
                st.toast("Calendar converted, scroll down to download", icon="🥳")
                st.balloons()



def process_xlsx(pbl, clin, comm, uni_format, campus, ws, preview, date_start, date_end):
    events = []
    saved = []
    rows_kept = []
    preview_error = False

    row_num = 1

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=20, values_only=True):
        events.append(row)

    for i in events:
        i = list(i)
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


def generate_cal(events, date_start, date_end, uni_format, previous=None):
    new_events = build_events(events, date_start, date_end, uni_format)

    if previous is None:
        show_events("Importing events...", new_events)
        with open("calendar.ics", "wb") as f:
            f.write(to_ics(new_events))
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
                built.append(event)
        except (AttributeError, ValueError) as e:
            print(e)
            if show_errors:
                st.error(f"Something went wrong. Check the spreadsheet for information on {i}\n{e}")
                st.toast(f"Something went wrong. Add {i} manually.", duration="infinite")

    return built


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
