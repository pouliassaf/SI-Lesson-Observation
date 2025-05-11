import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime
import os

# NOTE: Public deployment version with domain-restricted access
st.title("Weekly Lesson Observation Input Tool (Public)")

# Email gatekeeping
email = st.text_input("Enter your school email to continue")
allowed_domains = ["@charterschools.ae", "@adek.gov.ae"]

if not any(email.endswith(domain) for domain in allowed_domains):
    st.warning("Access restricted. Please use an authorized school email (e.g. @charterschools.ae or @adek.gov.ae).")
    st.stop()

uploaded_file = st.file_uploader("Upload your school's Excel workbook:", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    lo_sheets = [sheet for sheet in wb.sheetnames if sheet.startswith("LO ")]
    st.success(f"Found {len(lo_sheets)} LO sheets in workbook.")

    if "Guidelines" in wb.sheetnames:
        st.expander("📘 Click here to view observation guidelines").markdown(
            "\n".join([str(cell.value) for row in wb["Guidelines"].iter_rows(values_only=True) for cell in row if cell])
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

    st.write("### Observer and Observation Info")
    observer_name = st.text_input("Name of the Observer")
    teacher_name = st.text_input("Name of the Teacher Observed")
    operator_name = st.selectbox("Operator", sorted(["Taaleem", "Al Dar", "New Century Education", "Bloom"]))

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

    school_list = sorted(school_options.get(operator_name, []))
    school_name = st.selectbox("School Name", school_list)
    observation_type = st.selectbox("Type of Observation", ["Individual", "Joint"])
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ws["Z1"] = "Observer Name"
    ws["AA1"] = observer_name
    ws["Z2"] = "Teacher Observed"
    ws["AA2"] = teacher_name
    ws["Z3"] = "Observation Type"
    ws["AA3"] = observation_type
    ws["Z4"] = "Timestamp"
    ws["AA4"] = timestamp
    ws["Z5"] = "Operator"
    ws["AA5"] = operator_name
    ws["Z6"] = "School Name"
    ws["AA6"] = school_name
    ws["Z7"] = "Email"
    ws["AA7"] = email

    st.write("### General Lesson Info")
    grades = ["Grade " + str(i) for i in range(1, 13)] + ["K1", "K2"]
    grades.sort()

    b_inputs = {
        "Grade": st.selectbox("Grade", grades),
        "Date": st.date_input("Date"),
        "Gender": st.selectbox("Gender", ["Female", "Male", "Mixed"]),
        "Number of Students": st.text_input("Number of Students", value=ws["B6"].value or ""),
        "Number of Males": st.text_input("Number of Males", value=ws["B7"].value or ""),
        "Number of Females": st.text_input("Number of Females", value=ws["B8"].value or "")
    }

    subjects = sorted(["Math", "English", "Arabic", "Science", "Islamic", "Social Studies"])
    periods = [f"Period {i}" for i in range(1, 9)]

    d_inputs = {
        "Subject": st.selectbox("Subject", subjects),
        "Period": st.selectbox("Period", periods),
        "Time In": st.time_input("Time In (Start Time)"),
        "Time Out": st.time_input("Time Out (End Time)"),
        "Number of Students of Determination": st.text_input("Number of Students of Determination", value=ws["D5"].value or ""),
        "Number of Teaching Assistants": st.text_input("Number of Teaching Assistants", value=ws["D6"].value or "")
    }

    try:
        lesson_duration = datetime.combine(datetime.today(), d_inputs["Time Out"]) - datetime.combine(datetime.today(), d_inputs["Time In"])
        duration_minutes = round(lesson_duration.total_seconds() / 60)
        lesson_type = "Full Lesson" if duration_minutes >= 40 else "Walkthrough"
    except Exception:
        lesson_type = ""

    d_inputs["Lesson Duration"] = lesson_type
    st.markdown(f"**Auto-detected Lesson Duration:** {lesson_type}")

    st.write("---")
    st.write("### Rubric-Based Evaluation")

    domain_config = {
        1: (11, 15),
        2: (20, 22),
        3: (27, 30),
        4: (35, 37),
        5: (42, 43),
        6: (48, 49),
        7: (54, 55),
        8: (60, 62),
        9: (67, 68)
    }

    rubric_columns = ["C", "D", "E", "F", "G", "H"]

    for domain_num, (start_row, end_row) in domain_config.items():
        domain_label = ws[f"A{start_row}"].value or f"Domain {domain_num}"
        st.markdown(f"### {domain_label}")

        for row in range(start_row, end_row + 1):
            element = ws[f"B{row}"].value or f"Element at B{row}"
            rubric = [ws[f"{col}{row}"].value or "" for col in rubric_columns]
            rubric_text = "\n".join([f"**{i+1}** – {desc}" for i, desc in enumerate(rubric)])

            st.markdown(f"**{element}**")
            with st.expander("Rubric Guidance (1–6)"):
                st.markdown(rubric_text)

            score = st.number_input(f"Rating for '{element}' (1–6)", min_value=1, max_value=6, step=1, value=ws[f"I{row}"].value or 1, key=f"I{row}")
            ws[f"I{row}"] = score

    if "Observation Log" not in wb.sheetnames:
        log_ws = wb.create_sheet("Observation Log")
        log_ws.append(["Sheet", "Observer", "Teacher", "Operator", "School", "Type", "Email", "Timestamp"])
    else:
        log_ws: Worksheet = wb["Observation Log"]

    if st.button("Save Inputs"):
        ws["B2"] = school_name
        ws["B3"] = b_inputs["Grade"]
        ws["B4"] = b_inputs["Date"].strftime("%Y-%m-%d")
        ws["B5"] = b_inputs["Gender"]
        ws["B6"] = b_inputs["Number of Students"]
        ws["B7"] = b_inputs["Number of Males"]
        ws["B8"] = b_inputs["Number of Females"]

        ws["D2"] = d_inputs["Subject"]
        ws["D3"] = d_inputs["Lesson Duration"]
        ws["D4"] = d_inputs["Period"]
        ws["D5"] = d_inputs["Number of Students of Determination"]
        ws["D6"] = d_inputs["Number of Teaching Assistants"]
        ws["D7"] = d_inputs["Time In"].strftime("%H:%M")
        ws["D8"] = d_inputs["Time Out"].strftime("%H:%M")

        log_ws.append([sheet_name, observer_name, teacher_name, operator_name, school_name, observation_type, email, timestamp])

        save_path = f"updated_{uploaded_file.name}"
        wb.save(save_path)
        st.success(f"Workbook updated and saved as {save_path}")
        with open(save_path, "rb") as f:
            st.download_button("Download updated workbook", f, file_name=save_path)
        os.remove(save_path)
else:
    st.info("Please upload your Excel file to begin.")
