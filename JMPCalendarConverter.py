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

    university = st.radio("University", ["University of Newcastle", "University of New England (beta)"], horizontal=True)
    use_latest = False

    if university == "University of Newcastle":
        use_latest = st.checkbox("Use latest UON year 1 timetable (last updated: 18:34 25/07/2025)", value=True)

    elif university == "University of New England (beta)":
        use_latest = st.checkbox("Use latest UNE year 1 timetable (last updated: 12:28 25/07/2025)", value=True)

    if use_latest and university == "University of Newcastle":
        file = "UONMEDI1101B 25072025.xlsx"
    elif use_latest and university == "University of New England (beta)":
        file = "UNEMEDI1101B 25072025.xlsx"
    else:
        file = st.file_uploader("Upload your calendar file", type=None, accept_multiple_files=False)

    date_start = datetime.date(1970, 1, 1)
    date_end = datetime.date(3000, 1, 1)
    dates = []

    if file:
        # process file and determine which date ranges are valid
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        st.divider()
        st.subheader("Options")

        option = st.radio("Convert", ["All events", "Current week", "Custom dates"], horizontal=True)

        if option != "Custom dates":
            if option == "Suggested dates":
                dates = [datetime.date(2025, 7, 21), datetime.date(2025, 5, 22), datetime.date(2025, 6, 11)]
            elif option == "Current week":
                date_start = datetime.date(2025, 7, 27)
                date_end = datetime.date(2025, 8, 2)

        if university == "University of Newcastle":
            campus = st.selectbox("Campus", ["Callaghan", "Central Coast"])

            column1, column2 = st.columns(2)

            if option == "Custom dates":
                with column1:
                    date_start = st.date_input("Date to import from (inclusive)", value="today", min_value=None,
                                               max_value=None,
                                               format="DD/MM/YYYY", key="dsi")
                with column2:
                    date_end = st.date_input("Last date to import (inclusive)", value="today", min_value=None,
                                             max_value=None,
                                             format="DD/MM/YYYY", key="dei")

            with column1:
                pbl = st.selectbox("PBL group", (
                    "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"),
                                   index=None,
                                   placeholder="Select PBL group")
            with column2:
                if campus == "Central Coast":
                    clin = st.selectbox("Clinical group", (
                        "1", "2", "3", "4"), index=None, placeholder="Select clinical group")
                else:
                    clin = "5"


            if option == "All events":
                st.text(f"Importing all available events from {file}")
            elif option == "Suggested dates":
                st.text("Importing suggested updates: 21/05/2025-22/05/2025 and 11/06/2025")
            elif option == "Current week":
                st.text("Importing JMP week 17: 27/07/2025-2/08/2025")

            valid_selection = True

            if campus == "Callaghan" and pbl and pbl not in ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"]:
                st.error(f"Callaghan does not have PBL group {pbl}")
                clin = "5"
                valid_selection = False
            elif campus == "Central Coast" and pbl and clin:
                if pbl not in ["A", "B", "C", "D"]:
                    st.error(f"Central Coast does not have PBL group {pbl}")
                    valid_selection = False
                if clin not in ["1", "2", "3", "4"]:
                    st.error(f"Central Coast does not have clinical group {clin}")
                    valid_selection = False

            go = False

            col11, col12 = st.columns(2)

            if valid_selection:
                with col11:
                    go = st.button("Start converting", use_container_width=True)
            else:
                st.warning(f"Select your timetable.")

            if go and valid_selection:
                saved = process_xlsx(pbl.upper(), clin, campus, ws)
                generate_cal(saved, date_start, date_end, dates)

                if os.path.exists("calendar.ics"):
                    f = open("calendar.ics", "r")

                    with col12:
                        st.download_button("Download ics file", data=f, file_name="Calendar.ics", use_container_width=True)

                    st.text(
                        "Import this file to your calendar app (google calendar works idk about the rest)\nAlways double check to see if events have been imported correctly. DM me @photogrudesh on Instagram if there are any issues.")

        elif university == "University of New England (beta)":

            column1, column2 = st.columns(2)

            if option == "Custom dates":
                with column1:
                    date_start = st.date_input("Date to import from (inclusive)", value="today", min_value=None,
                                               max_value=None,
                                               format="DD/MM/YYYY", key="dsi")
                with column2:
                    date_end = st.date_input("Last date to import (inclusive)", value="today", min_value=None,
                                             max_value=None,
                                             format="DD/MM/YYYY", key="dei")

            pbl = st.selectbox("PBL group", ("A", "B", "C", "D", "E", "F", "G", "H", "I"),
                           index=None,
                           placeholder="Select PBL group")

            with column1:
                clin = st.selectbox("Clin group", ("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"),
                                   index=None,
                                   placeholder="Select clin group")
            with column2:
                comm = st.selectbox("Comm group", ("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"),
                                   index=None,
                                   placeholder="Select comm group")

            if option == "All events":
                st.text(f"Importing all available events from {file}")
            elif option == "Suggested dates":
                st.text("Importing suggested updates: 21/05/2025-22/05/2025 and 11/06/2025")
            elif option == "Current week":
                st.text("Importing JMP week 16: 21/07/2025-25/07/2025")

            go = False

            col11, col12 = st.columns(2)

            if pbl and clin and comm:
                with col11:
                    go = st.button("Start converting", use_container_width=True)
            else:
                st.warning(f"Select your timetable.")

            if go and pbl and clin and comm:
                saved = process_xlsx_une(pbl.upper(), clin, comm, ws)
                generate_cal_une(saved, date_start, date_end, dates)

                if os.path.exists("calendar.ics"):
                    f = open("calendar.ics", "r")

                    with col12:
                        st.download_button("Download ics file", data=f, file_name="Calendar.ics",
                                           use_container_width=True)

                    st.text(
                        "Import this file to your calendar app (google calendar works idk about the rest)\nAlways double check to see if events have been imported correctly. DM me @photogrudesh on Instagram if there are any issues.")


def process_xlsx_une(pbl, clin, comm, ws):
    title_row = 2
    titles = []

    for i in range(1, 15):
        titles.append(ws.cell(row=title_row, column=i).value)

    print(titles)

    # ['WEEK', 'DATE', 'DAY', 'TIME', 'DURATION', 'GROUPS', 'VENUE', 'TYPE', 'Attendance (M - Mandatory)', 'SESSION', 'PRESENTER', None, None, None]

    events = []
    saved = []

    row_num = title_row + 1

    for row in ws.iter_rows(min_row=title_row + 1, max_row=ws.max_row, max_col=12, values_only=True):
        events.append(row)

    for i in events:
        i = list(i)
        keep = False
        students = str(i[5])
        if "ALL" in students:
            keep = True
        elif students.startswith("PBL"):
            if bool(re.search(rf'(?<!\w){re.escape(pbl)}(?!\w)', students)):
                keep = True
        elif students.startswith("CLIN"):
            if bool(re.search(rf'(?<!\d){clin}(?!\d)', students)):
                keep = True
        elif students.startswith("COMM"):
            if bool(re.search(rf'(?<!\d){comm}(?!\d)', students)):
                keep = True
        elif i[7] == "Anatomy Lab":
            keep = True
        elif students == pbl:
            keep = True

        if "Zoom" in str(i[6]):
            try:
                i[7] = ws.cell(row=row_num, column=8).hyperlink.target
            except AttributeError:
                pass

        if keep:
            saved.append(i)
            print(i)
        row_num += 1

    return saved


def generate_cal_une(events, date_start, date_end, dates):
    cal = Calendar()
    cal['version'] = '2.0'
    cal.add("prodid", "-//photogrudesh//JMPCalendarConverter//EN")
    cal.add("summary", "JMP schedule")

    with st.status("Importing events..."):
        if not dates:
            for i in events:
                if date_start - datetime.timedelta(days=1) < i[1].date() < date_end + datetime.timedelta(days=1):
                    event = Event()
                    no_time = False

                    try:
                        start, end = convert_datetime(i[1], i[3])
                        event.add('dtstart', start)
                        event.add('dtend', end)
                    except TypeError:
                        no_time = True

                    if i[8] is None:
                        attendance = "N/A"
                    else:
                        attendance = i[8]

                    desc = f"{i[7]}\n{i[6]}\nStudents: {i[5]}\nAttendance: {attendance}\nStaff: {i[10]}\nUpdates: {i[11]}"

                    print(i[1])

                    if no_time:
                        event.add('dtstart', i[1])
                        event.add('dtend', i[1] + datetime.timedelta(days=1))

                    event.add('summary', i[9])
                    event.add("description", desc)
                    cal.add_component(event)
                    if not no_time:
                        st.write(f"{start}: " + i[9].replace('\n', ''))
                    else:
                        st.write(i[9].replace('\n', ''))

        elif dates:
            for i in events:
                for j in dates:
                    if i[3].date() == j:
                        event = Event()
                        no_time = False

                        try:
                            start, end = convert_datetime(i[1], i[3])
                            event.add('dtstart', start)
                            event.add('dtend', end)
                        except TypeError:
                            no_time = True

                        if i[8] is None:
                            attendance = "N/A"
                        else:
                            attendance = i[8]

                        desc = f"{i[7]}\n{i[6]}\nStudents: {i[5]}\nAttendance: {attendance}\nStaff: {i[10]}\nUpdates: {i[11]}"

                        if no_time:
                            event.add('dtstart', i[2])
                            event.add('dtend', i[2] + datetime.timedelta(days=1))

                        event.add('summary', i[9])
                        event.add("description", desc)

                        cal.add_component(event)
                        if not no_time:
                            st.write(f"{start}: " + i[9].replace('\n', ''))
                        else:
                            st.write(i[9].replace('\n', ''))

    with open("calendar.ics", "wb") as f:
        f.write(cal.to_ical())


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

        if "Zoom" in str(i[7]):
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


def generate_cal(events, date_start, date_end, dates):
    cal = Calendar()
    cal['version'] = '2.0'
    cal.add("prodid", "-//photogrudesh//JMPCalendarConverter//EN")
    cal.add("summary", "JMP schedule")

    with st.status("Importing events..."):

        if not dates:
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
                    if not no_time:
                        st.write(f"{start}: " + i[11].replace('\n', ''))
                    else:
                        st.write(i[11].replace('\n', ''))

        elif dates:
            for i in events:
                for j in dates:
                    if i[3].date() == j:
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
                        if not no_time:
                            st.write(f"{start}: " + i[11].replace('\n', ''))
                        else:
                            st.write(i[11].replace('\n', ''))

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
