import os

import streamlit as st
import openpyxl
from icalendar import Calendar, Event
import datetime
import parsedatetime as pdt
import re


def main():
    st.set_page_config(page_title="JMPCalendarConverter")

    st.title("JMPCalendarConverter")
    st.divider()
    st.subheader("Upload xlsx")

    use_latest = st.checkbox("Use latest year 1 timetable (last updated: 16:32 01/04/2025)", value=True)
    file = "MEDI1101 01042025.xlsx"

    date_start = datetime.date(1970, 1, 1)
    date_end = datetime.date(3000, 1, 1)

    suggested_start = "today"
    suggested_end = "today"

    if not use_latest:
        file = st.file_uploader("Upload your calendar file", type=None, accept_multiple_files=False)
    else:
        st.text("Suggested date range to update for Callaghan: 06/04/2025-12/04/2025")
        col1, col2 = st.columns(2)

        with col1:
            autofill = st.button("Apply suggested updates", use_container_width=True)
        with col2:
            current_week = st.button("Update current week", use_container_width=True)

        if autofill:
            suggested_start = datetime.date(2025, 4, 6)
            suggested_end = datetime.date(2025, 4, 12)
        if current_week:
            suggested_start = datetime.date(2025, 4, 6)
            suggested_end = datetime.date(2025, 4, 12)

    if file:
        # process file and determine which date ranges are valid
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        st.divider()
        st.subheader("Options")

        import_all = st.checkbox("Import all", key="importcheck")

        column1, column2 = st.columns(2)

        if not import_all:
            with column1:
                date_start = st.date_input("Date to import from (inclusive)", value=suggested_start, min_value=None,
                                           max_value=None,
                                           format="DD/MM/YYYY", key="dsi")
            with column2:
                date_end = st.date_input("Last date to import (inclusive)", value=suggested_end, min_value=None,
                                         max_value=None,
                                         format="DD/MM/YYYY", key="dei")

        with column1:
            pbl = st.selectbox("PBL group", (
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"),
                               index=None,
                               placeholder="Select PBL group")
        with column2:
            clin = st.selectbox("Clinical group", (
                "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19",
                "20"), index=None, placeholder="Select clinical group")

        campus = st.selectbox("Campus", ["Callaghan", "Central Coast"])

        valid_selection = True

        if pbl and clin and campus:
            if campus == "Callaghan":
                if pbl not in ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"]:
                    st.error(f"Callaghan does not have PBL group {pbl}")
                    valid_selection = False
                if clin not in ["5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19",
                                "20"]:
                    st.error(f"Callaghan does not have clinical group {clin}")
                    valid_selection = False
            elif campus == "Central Coast":
                if pbl not in ["A", "B", "C", "D"]:
                    st.error(f"Central Coast does not have PBL group {pbl}")
                    valid_selection = False
                if clin not in ["1", "2", "3", "4"]:
                    st.error(f"Central Coast does not have clinical group {clin}")
                    valid_selection = False

        go = False

        col11, col12 = st.columns(2)

        if valid_selection and pbl and clin and campus:
            with col11:
                go = st.button("Start converting", use_container_width=True)
        else:
            st.warning(f"Select your timetable.")

        if pbl and clin and campus and go and valid_selection:
            saved = process_xlsx(pbl.upper(), clin, campus, ws)
            generate_cal(saved, date_start, date_end)

            if os.path.exists("calendar.ics"):
                f = open("calendar.ics", "r")

                with col12:
                    st.download_button("Download ics file", data=f, file_name="Calendar.ics", use_container_width=True)

                st.text("Import this file to your calendar app (google calendar works idk about the rest)\nAlways double check to see if events have been imported correctly. DM me @photogrudesh on Instagram if there are any issues.")


def process_xlsx(pbl, clin, campus, ws):
    title_row = 3
    titles = []

    for i in range(1, 15):
        titles.append(ws.cell(row=title_row, column=i).value)

    # ['Campus', 'JMP Week', 'Day', 'Date', 'Time', 'Duration', 'Students', 'Venue', 'Type', 'Domain', 'Attendance\n(M =\nmandatory)', 'Name of Activity ', 'Staff', 'Update']

    events = []
    saved = []

    row_num = title_row + 1

    for row in ws.iter_rows(min_row=title_row + 1, max_row=ws.max_row, max_col=14, values_only=True):
        events.append(row)

    for i in events:
        i = list(i)
        keep = False
        students = str(i[6])
        if "ALL" in students:
            keep = True
        elif students.startswith("PBL"):
            if bool(re.search(rf'(?<!\w){re.escape(pbl)}(?!\w)', students)):
                keep = True
        elif students.startswith("CLIN"):
            if bool(re.search(rf'(?<!\d){clin}(?!\d)', students)):
                keep = True

        if i[7] == "Zoom":
            try:
                i[7] = ws.cell(row=row_num, column=8).hyperlink.target
            except AttributeError:
                pass

        if campus not in str(i[0]):
            keep = False

        if keep:
            saved.append(i)
        row_num += 1

    return saved


def generate_cal(events, date_start, date_end):
    cal = Calendar()
    cal.add("prodid", "-//photogrudesh//JMPCalendarConverter//EN")
    cal.add("version", "1.0.0")
    cal.add("summary", "JMP schedule")

    for i in events:
        if date_start - datetime.timedelta(days=1) < i[3].date() < date_end + datetime.timedelta(days=1):
            event = Event()
            no_time = False

            try:
                start, end = convert_datetime(i[3], i[4])
                event.add('dtstart', start)
                event.add('dtend', end)
            except TypeError:
                no_time = True

            if i[10] is None:
                attendance = "N/A"
            else:
                attendance = i[10]

            desc = f"{i[8]}: {i[9]}\n{i[7]}\nStudents: {i[6]}\nAttendance: {attendance}\nStaff: {i[12]}\nUpdates: {i[13]}"

            if no_time:
                event.add('dtstart', i[3])
                event.add('dtend', i[3] + datetime.timedelta(days=1))

            event.add('summary', i[11])
            event.add("description", desc)

            cal.add_component(event)

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
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
