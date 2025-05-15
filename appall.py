import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime
import os
import statistics

st.set_page_config(page_title="Lesson Observation Tool", layout="wide")

page = st.sidebar.selectbox("Choose a page:", ["Lesson Input", "Observation Analytics"])

if page == "Lesson Input":
    st.title("Weekly Lesson Observation Input Tool")

email = st.text_input("Enter your school email to continue")
allowed_domains = ["@charterschools.ae", "@adek.gov.ae"]

if not any(email.endswith(domain) for domain in allowed_domains):
    st.warning("Access restricted. Please use an authorized school email.")
    st.stop()

uploaded_file = None
DEFAULT_FILE = "Teaching Rubric Tool_WeekTemplate.xlsx"
if not uploaded_file and os.path.exists(DEFAULT_FILE):
    uploaded_file = open(DEFAULT_FILE, "rb")
    st.info("Using default template workbook.")

elif page == "Observation Analytics":
    import pandas as pd
    import matplotlib.pyplot as plt

    st.title("Observation Analytics")

    if os.path.exists(DEFAULT_FILE):
        wb = load_workbook(DEFAULT_FILE)
        if "Observation Log" in wb.sheetnames:
            ws = wb["Observation Log"]
            data = list(ws.values)
            headers, *rows = data
            df = pd.DataFrame(rows, columns=headers)
            st.success("Observation Log loaded successfully.")

            if not df.empty:
                st.subheader("Overall Judgment Distribution")
                st.bar_chart(df["Final Judgment"].value_counts())

                st.subheader("Average Score by School")
                avg_school = df.groupby("School")["Overall Avg"].mean().sort_values()
                st.bar_chart(avg_school)

                st.subheader("Average Score by Subject")
                if "Subject" in df.columns:
                    avg_subject = df.groupby("Subject")["Overall Avg"].mean().sort_values()
                    st.bar_chart(avg_subject)

                st.subheader("Average Score by Grade")
                if "Grade" in df.columns:
                    avg_grade = df.groupby("Grade")["Overall Avg"].mean().sort_values()
                    st.bar_chart(avg_grade)
            else:
                st.warning("Observation Log is empty.")
        else:
            st.warning("'Observation Log' sheet not found in the Excel file.")
    else:
        st.warning("Default file not found. Please upload or run a lesson observation first.")

if page == "Lesson Input" and uploaded_file:
    wb = load_workbook(uploaded_file)
    lo_sheets = [sheet for sheet in wb.sheetnames if sheet.startswith("LO ")]
    st.success(f"Found {len(lo_sheets)} LO sheets in workbook.")

    if st.checkbox("🪟 Clean up unused LO sheets (no observer name)"):
        to_remove = []
        for sheet in lo_sheets:
            if wb[sheet]["AA1"].value is None:
                to_remove.append(sheet)
        for sheet in to_remove:
            wb.remove(wb[sheet])
        st.warning(f"Removed {len(to_remove)} unused LO sheets.")
        lo_sheets = [sheet for sheet in wb.sheetnames if sheet.startswith("LO ")]

    if "Guidelines" in wb.sheetnames:
        st.expander("📘 Click here to view observation guidelines").markdown(
            "\n".join([str(cell) for row in wb["Guidelines"].iter_rows(values_only=True) for cell in row if cell])
        )

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
    operator = st.selectbox("Operator", sorted(["Taaleem", "Al Dar", "New Century Education", "Bloom"]))

    school_options = {
        "New Century Education": [
            "Al Bayan School", "Al Bayraq School", "Al Dhaher School", "Al Hosoon School",
            "Al Mutanabi School", "Al Nahdha School", "Jern Yafoor School", "Maryam Bint Omran School"
        ],
        "Taaleem": [...],
        "Al Dar": [...],
        "Bloom": [...]
    }

    school_list = sorted(school_options.get(operator, []))
    school = st.selectbox("School Name", school_list)
    grade = st.selectbox("Grade", [f"Grade {i}" for i in range(1, 13)] + ["K1", "K2"])
    date = st.date_input("Date")
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

    st.markdown("---")
    st.subheader("Rubric Scores")

    rubric_domains = {
        "Domain 1": ("I11", 5),
        "Domain 2": ("I20", 3),
        "Domain 3": ("I27", 4),
        "Domain 4": ("I35", 3),
        "Domain 5": ("I42", 2),
        "Domain 6": ("I48", 2),
        "Domain 7": ("I54", 2),
        "Domain 8": ("I60", 3),
        "Domain 9": ("I67", 2)
    }

    domain_colors = ["#e6f2ff", "#fff2e6", "#e6ffe6", "#f9e6ff", "#ffe6e6", "#f0f0f5", "#e6f9ff", "#f2ffe6", "#ffe6f2"]

    for idx, (domain, (start_cell, count)) in enumerate(rubric_domains.items()):
        background = domain_colors[idx % len(domain_colors)]
        st.markdown(f"<div style='background-color:{background};padding:12px;border-radius:10px;margin-bottom:5px;'><h4 style='margin-bottom:5px;'>{domain}: {ws[f'A{int(start_cell[1:])}'].value}</h4></div>", unsafe_allow_html=True)
        col = start_cell[0]
        row = int(start_cell[1:])
        scores = []
        for i in range(count):
            element_number = f"{idx+1}.{i+1}"
            label = ws[f"B{row + i}"].value or f"Element {element_number}"
            rubric = [ws[f"C{row + i}"].value, ws[f"D{row + i}"].value, ws[f"E{row + i}"].value, ws[f"F{row + i}"].value, ws[f"G{row + i}"].value, ws[f"H{row + i}"].value]
            formatted = "\n\n".join([f"**{6-j}:** {desc}" for j, desc in enumerate(rubric) if desc])
            st.markdown(f"<div style='background-color:{background};padding:8px;border-radius:6px;'>", unsafe_allow_html=True)
            st.markdown(f"**{element_number} – {label}**")
            with st.expander("Rubric Guidance"):
                st.markdown(formatted)
            col1, col2 = st.columns([1, 2])
            with col1:
                val = st.selectbox(f"Rating for {element_number}", options=[6, 5, 4, 3, 2, 1, "NA"], key=f"{domain}_{i}")
                ws[f"{col}{row + i}"] = val
                if isinstance(val, int):
                    scores.append(val)
            with col2:
                note = st.text_area(f"Write notes for {element_number}", key=f"{domain}_{i}_note", height=100)
                ws[f"J{row + i}"] = note
            st.markdown("</div>", unsafe_allow_html=True)

        if scores:
            avg = round(statistics.mean(scores), 2)
            if avg >= 5.5:
                judgment = "Outstanding"
            elif avg >= 4.5:
                judgment = "Very Good"
            elif avg >= 3.5:
                judgment = "Good"
            elif avg >= 2.5:
                judgment = "Acceptable"
            elif avg >= 1.5:
                judgment = "Weak"
            else:
                judgment = "Very Weak"
            st.success(f"✅ Domain Average: {avg} | Judgment: {judgment}")
        else:
            st.info("No numeric scores entered yet for this domain.")

        
    # Display overall average and judgment across all domains
    all_scores = []
    for domain, (start_cell, count) in rubric_domains.items():
        col = start_cell[0]
        row = int(start_cell[1:])
        for i in range(count):
            cell_value = ws[f"{col}{row + i}"].value
            if isinstance(cell_value, int):
                all_scores.append(cell_value)

    if all_scores:
        overall_avg = round(statistics.mean(all_scores), 2)
        if overall_avg >= 5.5:
            overall_judgment = "Outstanding"
        elif overall_avg >= 4.5:
            overall_judgment = "Very Good"
        elif overall_avg >= 3.5:
            overall_judgment = "Good"
        elif overall_avg >= 2.5:
            overall_judgment = "Acceptable"
        elif overall_avg >= 1.5:
            overall_judgment = "Weak"
        else:
            overall_judgment = "Very Weak"

        st.markdown("---")
        st.success(f"🌟 Overall Average: {overall_avg} | Final Judgment: {overall_judgment}")
    else:
        st.info("No numeric scores entered yet for overall calculation.")

    overall_notes = st.text_area("General Notes for this Lesson Observation")

        # Remaining unchanged code for saving to Excel, etc.

    if st.button("💾 Save this Observation"):
        for domain_label, (start_cell, count) in rubric_domains.items():
            col = start_cell[0]
            row = int(start_cell[1:])
            score_range = f"{col}{row}:{col}{row + count - 1}"
            avg_cell = f"{col}{row + count}"
            judgment_cell = f"{col}{row + count + 1}"
            ws[avg_cell] = f'=IF(COUNTA({score_range})=0, "", AVERAGEIF({score_range}, "<>NA"))'
            ws[judgment_cell] = f'=IF({avg_cell}="", "", IF({avg_cell}>=5.5,"Outstanding",IF({avg_cell}>=4.5,"Very Good",IF({avg_cell}>=3.5,"Good",IF({avg_cell}>=2.5,"Acceptable",IF({avg_cell}>=1.5,"Weak","Very Weak"))))))'

        ws["B5"] = gender
        ws["B6"] = students
        ws["B7"] = males
        ws["B8"] = females
        ws["D2"] = subject
        ws["D3"] = duration_label
        ws["D4"] = period
        ws["D7"] = time_in.strftime("%H:%M")
        ws["D8"] = time_out.strftime("%H:%M")

        ws["Z1"] = "Observer Name"
        ws["AA1"] = observer
        ws["Z2"] = "Teacher Observed"
        ws["AA2"] = teacher
        ws["Z3"] = "Observation Type"
        ws["AA3"] = obs_type
        ws["Z4"] = "Timestamp"
        ws["AA4"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws["Z5"] = "Operator"
        ws["AA5"] = operator
        ws["Z6"] = "School Name"
        ws["AA6"] = school
        ws["Z7"] = "General Notes"
        ws["AA7"] = overall_notes
        ws["Z8"] = "Overall Average"
        ws["AA8"] = overall_avg if all_scores else ""
        ws["Z9"] = "Final Judgment"
        ws["AA9"] = overall_judgment if all_scores else ""

        if "Observation Log" not in wb.sheetnames:
            log_ws = wb.create_sheet("Observation Log")
            log_ws.append(["Sheet", "Observer", "Teacher", "Operator", "School", "Type", "Timestamp", "Notes", "Overall Avg", "Final Judgment"])
        else:
            log_ws: Worksheet = wb["Observation Log"]

        log_ws.append([sheet_name, observer, teacher, operator, school, obs_type, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), overall_notes, overall_avg if all_scores else "", overall_judgment if all_scores else ""])

        save_path = f"updated_{sheet_name}.xlsx"
        wb.save(save_path)
        with open(save_path, "rb") as f:
            st.download_button("📅 Download updated workbook", f, file_name=save_path)
        os.remove(save_path)

















  

  










