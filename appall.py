import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime
import os

st.title("Weekly Lesson Observation Input Tool")

email = st.text_input("Enter your school email to continue")
allowed_domains = ["@charterschools.ae", "@adek.gov.ae"]

if not any(email.endswith(domain) for domain in allowed_domains):
    st.warning("Access restricted. Please use an authorized school email.")
    st.stop()

# uploaded_file = st.file_uploader("Upload your school's Excel workbook:", type=["xlsx"])
uploaded_file = None

DEFAULT_FILE = "Teaching Rubric Tool_WeekTemplate.xlsx"
if not uploaded_file and os.path.exists(DEFAULT_FILE):
    uploaded_file = open(DEFAULT_FILE, "rb")
    st.info("Using default template workbook.")

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    lo_sheets = [sheet for sheet in wb.sheetnames if sheet.startswith("LO ")]
    st.success(f"Found {len(lo_sheets)} LO sheets in workbook.")

    if st.checkbox("🧹 Clean up unused LO sheets (no observer name)"):
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

    st.text_input("Observer Name", key="observer")
    st.text_input("Teacher Name", key="teacher")
    st.selectbox("Operator", sorted(["Taaleem", "Al Dar", "New Century Education", "Bloom"]), key="operator")

    school_options = {
        "New Century Education": [
            "Al Bayan School", "Al Bayraq School", "Al Dhaher School", "Al Hosoon School",
            "Al Mutanabi School", "Al Nahdha School", "Jern Yafoor School", "Maryam Bint Omran School"
        ],
        "Taaleem": [
            "Al Ahad Charter School", "Al Azm Charter School", "Al Riyadh Charter School", "Al Majd Charter School",
            "Al Qeyam Charter School", "Al Nayfa Charter Kindergarten", "Al Salam Charter School",
            "Al Walaa Charter Kindergarten", "Al Forsan Charter Kindergarten", "Al Wafaa Charter Kindergarten",
            "Al Watan Charter School"
        ],
        "Al Dar": [
            "Al Ghad Charter School", "Al Mushrif Charter Kindergarten", "Al Danah Charter School",
            "Al Rayaheen Charter School", "Al Rayana Charter School", "Al Qurm Charter School",
            "Mubarak Bin Mohammed Charter School (Cycle 2 & 3)"
        ],
        "Bloom": [
            "Al Ain Charter School", "Al Dana Charter School", "Al Ghadeer Charter School", "Al Hili Charter School",
            "Al Manhal Charter School", "Al Qattara Charter School", "Al Towayya Charter School",
            "Jabel Hafeet Charter School"
        ]
    }

    school_list = sorted(school_options.get(st.session_state.operator, []))
    st.selectbox("School Name", school_list, key="school")
    st.selectbox("Grade", [f"Grade {i}" for i in range(1, 13)] + ["K1", "K2"], key="grade")
    st.date_input("Date", key="date")
    st.selectbox("Subject", ["Math", "English", "Arabic", "Science", "Islamic", "Social Studies"], key="subject")
    st.selectbox("Gender", ["Male", "Female", "Mixed"], key="gender")
    st.text_input("Number of Students", key="students")
    st.text_input("Number of Males", key="males")
    st.text_input("Number of Females", key="females")
    st.time_input("Time In", key="in")
    st.time_input("Time Out", key="out")
    st.selectbox("Period", [f"Period {i}" for i in range(1, 9)], key="period")
    st.selectbox("Observation Type", ["Individual", "Joint"], key="obs_type")

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

    for domain, (start_cell, count) in rubric_domains.items():
        st.markdown(f"**{domain}**")
        col = start_cell[0]
        row = int(start_cell[1:])
        for i in range(count):
            label = ws[f"B{row + i}"].value or f"Element {i+1}"
            rubric = [ws[f"C{row + i}"].value, ws[f"D{row + i}"].value, ws[f"E{row + i}"].value,
                      ws[f"F{row + i}"].value, ws[f"G{row + i}"].value, ws[f"H{row + i}"].value]
            tooltip = "
".join([f"{j+1}: {desc}" for j, desc in enumerate(rubric) if desc])
".join([f"{j+1}: {desc}" for j, desc in enumerate(rubric) if desc])"""
            val = st.number_input(f"{label}", min_value=1, max_value=6, key=f"{domain}_{i}", help=tooltip)
                        ws[f"{col}{row + i}"] = val

    if st.button("Save this Observation"):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws["AA1"] = st.session_state.observer
        ws["AA2"] = st.session_state.teacher
        ws["AA3"] = st.session_state.obs_type
        ws["AA5"] = st.session_state.operator
        ws["AA6"] = st.session_state.school
        ws["Z4"] = now
        ws["Z7"] = email

        ws["B3"] = st.session_state.grade
        ws["B4"] = st.session_state.date.strftime("%Y-%m-%d")
        ws["D2"] = st.session_state.subject
        ws["B5"] = st.session_state.gender
        ws["B6"] = st.session_state.students
        ws["B7"] = st.session_state.males
        ws["B8"] = st.session_state.females
        ws["D7"] = st.session_state["in"].strftime("%H:%M")
        ws["D8"] = st.session_state["out"].strftime("%H:%M")
        ws["D4"] = st.session_state.period

        filename = f"updated_{sheet_name}.xlsx"
        wb.save(filename)
        with open(filename, "rb") as f:
            st.download_button("📥 Download updated workbook", f, file_name=filename)
        os.remove(filename)






