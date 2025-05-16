import streamlit as st
from openpyxl import load_workbook
from datetime import datetime
import os
import statistics
import pandas as pd
import matplotlib.pyplot as plt
import csv

st.set_page_config(page_title="Lesson Observation Tool", layout="wide")

uploaded_file = None
DEFAULT_FILE = "Teaching Rubric Tool_WeekTemplate.xlsx"
if not uploaded_file and os.path.exists(DEFAULT_FILE):
    uploaded_file = open(DEFAULT_FILE, "rb")
    st.info("Using default template workbook.")

page = st.sidebar.selectbox("Choose a page:", ["Lesson Input", "Observation Analytics"])

if page == "Lesson Input":
    st.title("Weekly Lesson Observation Input Tool")

    st.markdown("""
    <style>
    .block-container {
        padding-top: 2rem;
    }
    </style>
    """, unsafe_allow_html=True)

    if uploaded_file:
        wb = load_workbook(uploaded_file)
        lo_sheets = [sheet for sheet in wb.sheetnames if sheet.startswith("LO ")]
        st.success(f"Found {len(lo_sheets)} LO sheets in workbook.")

        selected_option = st.selectbox("Select existing LO sheet or create a new one:", ["Create new"] + lo_sheets)

        if selected_option == "Create new":
            next_index = 1
            while f"LO {next_index}" in wb.sheetnames:
                next_index += 1
            sheet_name = f"LO {next_index}"
            wb.copy_worksheet(wb["LO 1"]).title = sheet_name
            st.success(f"Created new sheet: {sheet_name}")
        else:
            sheet_name = selected_option

        ws = wb[sheet_name]
        st.subheader(f"Filling data for: {sheet_name}")

        observer = st.text_input("Observer Name")
        teacher = st.text_input("Teacher Name")
        teacher_email = st.text_input("Teacher Email")
        operator = st.selectbox("Operator", sorted(["Taaleem", "Al Dar", "New Century Education", "Bloom"]))

        school_options = {
            "New Century Education": ["Al Bayan School", "Al Bayraq School", "Al Dhaher School", "Al Hosoon School", "Al Mutanabi School", "Al Nahdha School", "Jern Yafoor School", "Maryam Bint Omran School"],
            "Taaleem": ["Al Ahad Charter School", "Al Azm Charter School", "Al Riyadh Charter School", "Al Majd Charter School", "Al Qeyam Charter School", "Al Nayfa Charter Kindergarten", "Al Salam Charter School", "Al Walaa Charter Kindergarten", "Al Forsan Charter Kindergarten", "Al Wafaa Charter Kindergarten", "Al Watan Charter School"],
            "Al Dar": ["Al Ghad Charter School", "Al Mushrif Charter Kindergarten", "Al Danah Charter School", "Al Rayaheen Charter School", "Al Rayana Charter School", "Al Qurm Charter School", "Mubarak Bin Mohammed Charter School (Cycle 2 & 3)"],
            "Bloom": ["Al Ain Charter School", "Al Dana Charter School", "Al Ghadeer Charter School", "Al Hili Charter School", "Al Manhal Charter School", "Al Qattara Charter School", "Al Towayya Charter School", "Jabel Hafeet Charter School"]
        }

        school_list = sorted(school_options.get(operator, []))
        school = st.selectbox("School Name", school_list)
        grade = st.selectbox("Grade", [f"Grade {i}" for i in range(1, 13)] + ["K1", "K2"])
        subject = st.selectbox("Subject", ["Math", "English", "Arabic", "Science", "Islamic", "Social Studies"])
        gender = st.selectbox("Gender", ["Male", "Female", "Mixed"])
        students = st.text_input("Number of Students")
        males = st.text_input("Number of Males")
        females = st.text_input("Number of Females")
        time_in = st.time_input("Time In")
        time_out = st.time_input("Time Out")

        try:
            lesson_duration = datetime.combine(datetime.today(), time_out) - datetime.combine(datetime.today(), time_in)
            minutes = round(lesson_duration.total_seconds() / 60)
            duration_label = "Full Lesson" if minutes >= 40 else "Walkthrough"
            st.markdown(f"🕒 **Lesson Duration:** {minutes} minutes — _{duration_label}_")
        except Exception:
            st.warning("Could not calculate lesson duration.")

        period = st.selectbox("Period", [f"Period {i}" for i in range(1, 9)])
        obs_type = st.selectbox("Observation Type", ["Individual", "Joint"])

        rubric_domains = {
            "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
            "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
        }

        st.markdown("---")
        st.subheader("Rubric Scores")

        domain_colors = ["#e6f2ff", "#fff2e6", "#e6ffe6", "#f9e6ff", "#ffe6e6", "#f0f0f5", "#e6f9ff", "#f2ffe6", "#ffe6f2"]
        for idx, (domain, (start_cell, count)) in enumerate(rubric_domains.items()):
            background = domain_colors[idx % len(domain_colors)]
            st.markdown(f"""
            <div style='background-color:{background};padding:12px;border-radius:10px;margin-bottom:5px;'>
            <h4 style='margin-bottom:5px;'>{domain}: {ws[f'A{int(start_cell[1:])}'].value}</h4>
            </div>
            """, unsafe_allow_html=True)
            col = start_cell[0]
            row = int(start_cell[1:])
            for i in range(count):
                label = ws[f"B{row + i}"].value or f"Element {domain[-1]}.{i+1}"
                rubric = [ws[f"C{row + i}"].value, ws[f"D{row + i}"].value, ws[f"E{row + i}"].value, ws[f"F{row + i}"].value, ws[f"G{row + i}"].value, ws[f"H{row + i}"].value]
                rubric_text = "\n\n".join([f"**{6-j}:** {desc}" for j, desc in enumerate(rubric) if desc])
                st.markdown(f"**{label}**")
                with st.expander("Rubric Descriptors"):
                    st.markdown(rubric_text)
                rating = st.selectbox(f"Rating for {label}", [6, 5, 4, 3, 2, 1, "NA"], key=f"{sheet_name}_{domain}_{i}")
                ws[f"{col}{row + i}"] = rating

        send_feedback = st.checkbox("✉️ Send Feedback to Teacher")

        if st.button("💾 Save Observation"):
            ws["Z1"] = "Observer Name"; ws["AA1"] = observer
            ws["Z2"] = "Teacher"; ws["AA2"] = teacher
            ws["Z3"] = "Observation Type"; ws["AA3"] = obs_type
            ws["Z4"] = "Operator"; ws["AA4"] = operator
            ws["Z5"] = "School"; ws["AA5"] = school
            ws["Z6"] = "Subject"; ws["AA6"] = subject
            ws["Z7"] = "Grade"; ws["AA7"] = grade
            ws["Z8"] = "Gender"; ws["AA8"] = gender
            ws["Z9"] = "Students"; ws["AA9"] = students
            ws["Z10"] = "Males"; ws["AA10"] = males
            ws["Z11"] = "Females"; ws["AA11"] = females
            ws["Z12"] = "Duration"; ws["AA12"] = duration_label
            ws["Z13"] = "Time In"; ws["AA13"] = time_in.strftime("%H:%M")
            ws["Z14"] = "Time Out"; ws["AA14"] = time_out.strftime("%H:%M")

            save_path = f"updated_{sheet_name}.xlsx"
            wb.save(save_path)
            with open(save_path, "rb") as f:
                st.download_button("📥 Download updated workbook", f, file_name=save_path)
            os.remove(save_path)

            if send_feedback and teacher_email:
                feedback = (
                    f"Dear {teacher},\n\n"
                    "Your lesson observation has been saved.\n"
                    f"Observer: {observer}\n"
                    f"Duration: {duration_label}\n"
                    f"Subject: {subject}\n"
                    f"School: {school}\n\n"
                    "Based on rubric ratings, please review your updated workbook for details.\n\n"
                    "Regards,\nSchool Leadership Team"
                )
                st.success("Feedback generated (simulated):\n\n" + feedback)

                # Feedback log to sheet
                if "Feedback Log" not in wb.sheetnames:
                    log_ws = wb.create_sheet("Feedback Log")
                    log_ws.append(["Sheet", "Teacher", "Email", "Observer", "School", "Subject", "Date", "Summary"])
                else:
                    log_ws = wb["Feedback Log"]
                log_ws.append([
                    sheet_name, teacher, teacher_email, observer, school, subject,
                    datetime.now().strftime("%Y-%m-%d %H:%M"), feedback[:100] + ("..." if len(feedback) > 100 else "")
                ])

                # Feedback log as CSV
                log_csv_path = "feedback_log.csv"
                with open(log_csv_path, "w", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Sheet", "Teacher", "Email", "Observer", "School", "Subject", "Date", "Summary"])
                    for row in log_ws.iter_rows(min_row=2, values_only=True):
                        writer.writerow(row)
                with open(log_csv_path, "rb") as f:
                    st.download_button("📥 Download Feedback Log (CSV)", f, file_name=log_csv_path)
                os.remove(log_csv_path)
