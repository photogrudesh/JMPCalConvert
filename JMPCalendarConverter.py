import os

import streamlit as st
import openpyxl
from icalendar import Calendar, Event
import datetime
import parsedatetime as pdt
import re


def main_ui():
    st.set_page_config(page_title="JMPCalendarConverter")

    st.title("JMPCalendarConverter")
    st.text("Works well enough for now, I'm still ironing out bugs. Lmk if you notice anything @photogrudesh.")

    st.divider()

    university = st.radio("University - Year 2 only rn", ["University of Newcastle", "University of New England", "Use your own file"], horizontal=True)
    uni_format = None
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

    if university == "University of Newcastle":
        file = st.selectbox("Select a UON timetable", uoncals, placeholder="Select timetable")
        uni_format = "University of Newcastle"

    if university == "University of New England":
        file = st.selectbox("Select a UNE timetable", unecals, placeholder="Select timetable")
        uni_format = "University of New England"


    if university == "Use your own file":
        file = st.file_uploader("Upload your own calendar file", type=None, accept_multiple_files=False)
        uni_format = st.radio("Format",
                              ["University of Newcastle", "University of New England"], horizontal=True)

    dates = []
    date_start, date_end = None, None

    valid_selection = False

    if file:
        # process file and determine which date ranges are valid
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        st.divider()
        st.subheader("Options")

        option = st.radio("Convert", ["All events", "Custom dates"], horizontal=True)

        campus = None
        pbl = None

        column1, column2 = st.columns(2)

        if option == "All events":
            date_start = datetime.date(1970, 1, 1)
            date_end = datetime.date(3000, 1, 1)
        elif option == "Custom dates":
            with column1:
                date_start = st.date_input("Date to import from (inclusive)", value="today", min_value=None,
                                           max_value=None,
                                           format="DD/MM/YYYY", key="dsi")
            with column2:
                date_end = st.date_input("Last date to import (inclusive)", value="today", min_value=None,
                                         max_value=None,
                                         format="DD/MM/YYYY", key="dei")


        if uni_format == "University of Newcastle":
            campus = st.selectbox("Campus", ["Callaghan", "Central Coast"])

            pbl = st.selectbox("PBL group", (
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"),
                               index=None,
                               placeholder="Select PBL group")

            valid_selection = True

            if campus == "Callaghan" and pbl and pbl not in ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"]:
                st.error(f"Callaghan does not have PBL group {pbl}")
                valid_selection = False
            elif campus == "Central Coast" and pbl and pbl not in ["A", "B", "C", "D"]:
                    st.error(f"Central Coast does not have PBL group {pbl}")
                    valid_selection = False
            clin, comm = "12", "Z"

            go = False

        elif uni_format == "University of New England":

            pbl = st.selectbox("PBL group", ("A", "B", "C", "D", "E", "F", "G", "H", "I"),
                           index=None,
                           placeholder="Select PBL group")

            colclin, colcomm = st.columns(2)

            with colclin:
                clin = st.selectbox("Clin group", ("1", "2", "3", "4", "5", "6", "7", "8", "9", "10"), index=None, placeholder="Select clin group")
            with colcomm:
                comm = st.selectbox("Comm group", ("A", "B", "C", "D", "E", "F", "G", "H", "I"), index=None, placeholder="Select comm group")

            if option == "All events":
                st.text(f"Importing all available events from {file}")
            elif option == "Suggested dates":
                st.text("Importing suggested updates: 21/05/2025-22/05/2025 and 11/06/2025")
            elif option == "Current week":
                st.text("Importing JMP week 17: 27/07/2025-2/08/2025")

            if clin and comm and pbl:
                valid_selection = True

        go = False


        if valid_selection and pbl:
            go = st.button("Start converting", use_container_width=True)
            st.text("Always check the output below for errors and double check calendar events with the spreadsheet on canvas. Scroll down to download the file when it's ready.")
        else:
            st.warning(f"Select your timetable.")

        if go and valid_selection:
            saved = process_xlsx(pbl.upper(), clin, comm, uni_format, campus, ws)
            generate_cal(saved, date_start, date_end, dates, uni_format)

            if os.path.exists("calendar.ics"):
                f = open("calendar.ics", "r")

                st.download_button("Download ics file", data=f, file_name="Calendar.ics", use_container_width=True)

                st.text(
                    "Import this file to your calendar app (google calendar works idk about the rest)\nAlways double check to see if events have been imported correctly. DM me @photogrudesh on Instagram if there are any issues.")


def cal_index(campus, item):
    if campus == "University of Newcastle":
        match item:
            case "week":
                return 0
            case "day":
                return 1
            case "date":
                return 2
            case "time":
                return 3
            case "campus":
                return 5
            case "group":
                return 6
            case "venue":
                return 7
            case "attendance":
                return 8
            case "type":
                return 9
            case "domain":
                return 10
            case "session":
                return 11
            case "staff":
                return 12
            case "updates":
                return 13
    elif campus == "University of New England":
        match item:
            case "week":
                return 0
            case "day":
                return 1
            case "date":
                return 2
            case "time":
                return 3
            case "campus":
                return 6
            case "group":
                return 5
            case "venue":
                return 6
            case "attendance":
                return 10
            case "type":
                return 9
            case "domain":
                return 8
            case "session":
                return 7
            case "staff":
                return 11
            case "updates":
                return 12
    return None


def process_xlsx(pbl, clin, comm, uni_format, campus, ws):
    title_row = 1
    titles = []

    for i in range(1, 15):
        titles.append(ws.cell(row=title_row, column=i).value)

    events = []
    saved = []

    row_num = title_row + 1

    for row in ws.iter_rows(min_row=title_row + 1, max_row=ws.max_row, max_col=14, values_only=True):
        events.append(row)

    for i in events:
        i = list(i)
        keep = False
        group = str(i[cal_index(uni_format, "group")]).strip()
        if "ALL" in group:
            keep = True
        elif group.startswith("PBL"):
            if bool(re.search(rf'(?<!\w){re.escape(pbl)}(?!\w)', group)):
                keep = True
            if "-" in group:
                start_char = group[group.index("-") - 1]
                end_char = group[group.index("-") + 1]
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
            if "-" in group:
                start, end = group.split("-")
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
            if "-" in group:
                start_char = group[group.index("-") - 1]
                end_char = group[group.index("-") + 1]
                for x in [chr(c) for c in range(ord(start_char), ord(end_char) + 1)]:
                    if pbl == str(x):
                        keep = True
            elif "," in group:
                comms = group.split(",")
                for f in comms:
                    if f.strip() == comm:
                        keep = True
        if "zoom" in str(i[cal_index(uni_format, "venue")]).lower():

            zoom_col = 8
            try:
                if uni_format == "University of New England":
                    zoom_col = 7
                i[cal_index(uni_format, "venue")] = ws.cell(row=row_num, column=zoom_col).hyperlink.target
                print(i[cal_index(uni_format, "venue")])
            except AttributeError:
                print(f"No link for {i}")
                pass

        if uni_format == "University of Newcastle":
            campus_code = None
            if campus == "Callaghan":
                campus_code = "CAL"
            elif campus == "Central Coast":
                campus_code = "CC"
            if campus_code not in str(i[cal_index(uni_format, "campus")]):
                keep = False

        if keep:
            saved.append(i)
        row_num += 1
    return saved


def generate_cal(events, date_start, date_end, dates, uni_format):
    cal = Calendar()
    cal['version'] = '2.0'
    cal.add("prodid", "-//photogrudesh//JMPCalendarConverter//EN")
    cal.add("summary", "JMP schedule")

    with st.status("Importing events...", expanded=True):
            if not dates:
                for i in events:
                    try:
                        if date_start - datetime.timedelta(days=1) < i[cal_index(uni_format, "date")].date() < date_end + datetime.timedelta(days=1):
                            event = Event()
                            no_time = False

                            try:
                                start, end = convert_datetime(i[cal_index(uni_format, "date")], i[cal_index(uni_format, "time")])
                                event.add('dtstart', start)
                                event.add('dtend', end)
                            except TypeError:
                                no_time = True

                            if i[cal_index(uni_format, "attendance")] is None:
                                attendance = "N/A"
                            else:
                                attendance = i[cal_index(uni_format, "attendance")]

                            desc = f"{i[cal_index(uni_format, "type")]}: {i[cal_index(uni_format, "domain")]}\n{i[cal_index(uni_format, "venue")]}\nStudents: {i[cal_index(uni_format, "group")]}\nAttendance: {attendance}\nStaff: {i[cal_index(uni_format, "staff")]}\nUpdates: {i[cal_index(uni_format, "updates")]}"

                            if no_time:
                                event.add('dtstart', i[cal_index(uni_format, "date")])
                                event.add('dtend', i[cal_index(uni_format, "date")] + datetime.timedelta(days=1))

                            event.add('summary', i[cal_index(uni_format, "session")])
                            event.add("description", desc)
                            cal.add_component(event)
                            if not no_time:
                                st.write(f"{start}: " + i[cal_index(uni_format, "session")].replace('\n', ' '))
                            else:
                                st.write(i[cal_index(uni_format, "session")].replace('\n', ' '))
                    except AttributeError:
                        st.error(f"Something went wrong. Check the spreadsheet for information on {i}")
            elif dates:
                for i in events:
                    for j in dates:
                        try:
                            if i[cal_index(uni_format, "date")].date() == j:
                                event = Event()
                                no_time = False

                                try:
                                    start, end = convert_datetime(i[cal_index(uni_format, "date")],
                                                                  i[cal_index(uni_format, "time")])
                                    event.add('dtstart', start)
                                    event.add('dtend', end)
                                except TypeError:
                                    no_time = True

                                if i[cal_index(uni_format, "attendance")] is None:
                                    attendance = "N/A"
                                else:
                                    attendance = i[cal_index(uni_format, "attendance")]

                                desc = f"{i[cal_index(uni_format, "type")]}: {i[cal_index(uni_format, "domain")]}\n{i[cal_index(uni_format, "venue")]}\nStudents: {i[cal_index(uni_format, "group")]}\nAttendance: {attendance}\nStaff: {i[cal_index(uni_format, "staff")]}\nUpdates: {i[cal_index(uni_format, "updates")]}"

                                if no_time:
                                    event.add('dtstart', i[cal_index(uni_format, "date")])
                                    event.add('dtend', i[cal_index(uni_format, "date")] + datetime.timedelta(days=1))

                                event.add('summary', i[cal_index(uni_format, "session")])
                                event.add("description", desc)
                                cal.add_component(event)
                                if not no_time:
                                    st.write(f"{start}: " + i[cal_index(uni_format, "session")].replace('\n', ' '))
                                else:
                                    st.write(i[cal_index(uni_format, "session")].replace('\n', ' '))
                        except AttributeError:
                            st.error(f"Something went wrong. Check the spreadsheet for information on {i}")
    with open("calendar.ics", "wb") as f:
        f.write(cal.to_ical())


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
