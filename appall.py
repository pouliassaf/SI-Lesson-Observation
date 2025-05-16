import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime
import os
import statistics

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

st.markdown("---")
st.subheader("Rubric Scores")

for domain, (start_cell, count) in rubric_domains.items():
    col = start_cell[0]
    row = int(start_cell[1:])
    st.markdown(f"**{domain}**: {ws[f'A{row}'].value}")
    for i in range(count):
        label = ws[f"B{row + i}"].value or f"Element {domain[-1]}.{i+1}"
        rating = st.selectbox(f"Rating for {label}", [6, 5, 4, 3, 2, 1, "NA"], key=f"{domain}_{i}")
        ws[f"{col}{row + i}"] = rating

send_feedback = st.checkbox("✉️ Send Feedback to Teacher")

if st.button("💾 Save Observation"):
            ws["Z1"] = "Observer Name"
            ws["AA1"] = observer
            ws["Z2"] = "Teacher"
            ws["AA2"] = teacher
            ws["Z3"] = "Observation Type"
            ws["AA3"] = obs_type
            ws["Z4"] = "Operator"
            ws["AA4"] = operator
            ws["Z5"] = "School"
            ws["AA5"] = school
            ws["Z6"] = "Subject"
            ws["AA6"] = subject
            ws["Z7"] = "Grade"
            ws["AA7"] = grade
            ws["Z8"] = "Gender"
            ws["AA8"] = gender
            ws["Z9"] = "Students"
            ws["AA9"] = students
            ws["Z10"] = "Males"
            ws["AA10"] = males
            ws["Z11"] = "Females"
            ws["AA11"] = females
            ws["Z12"] = "Duration"
            ws["AA12"] = duration_label
            ws["Z13"] = "Time In"
            ws["AA13"] = time_in.strftime("%H:%M")
            ws["Z14"] = "Time Out"
            ws["AA14"] = time_out.strftime("%H:%M")

            save_path = f"updated_{sheet_name}.xlsx"
            wb.save(save_path)
            with open(save_path, "rb") as f:
                st.download_button("📥 Download updated workbook", f, file_name=save_path)
            os.remove(save_path)

            if send_feedback and teacher_email:
    feedback = (
        f"Dear {teacher},

"
        "Your lesson observation has been saved.
"
        f"Observer: {observer}
"
        f"Duration: {duration_label}
"
        f"Subject: {subject}
"
        f"School: {school}

"
        "Based on rubric ratings, please review your updated workbook for details.

"
        "Regards,
School Leadership Team"
    )
    st.success("Feedback generated (simulated):

" + feedback):

{feedback}"):

{feedback}"):

{feedback}"):

{feedback}")


st.success(f"Feedback generated for {teacher_email} (not sent, simulated):

{feedback}")

elif page == "Observation Analytics":
    st.title("Observation Analytics Dashboard")

    if uploaded_file:
        wb = load_workbook(uploaded_file, data_only=True)
        sheets = [s for s in wb.sheetnames if s.startswith("LO ")]

        domain_scores = {domain: [] for domain in [f"Domain {i}" for i in range(1, 10)]}
        metadata = []

        for sheet in sheets:
            ws = wb[sheet]
            row_info = {
                "School": ws["AA5"].value,
                "Grade": ws["AA7"].value,
                "Subject": ws["AA6"].value,
                "Observer": ws["AA1"].value
            }
            metadata.append(row_info)
            for i, (domain, (start_cell, count)) in enumerate({
                "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
            }.items()):
                col = start_cell[0]
                row = int(start_cell[1:])
                ratings = [ws[f"{col}{row + j}"].value for j in range(count) if isinstance(ws[f"{col}{row + j}"].value, int)]
                if ratings:
                    domain_scores[domain].append(statistics.mean(ratings))

        import pandas as pd
        import matplotlib.pyplot as plt

        avg_scores = {domain: round(statistics.mean(scores), 2) if scores else 0 for domain, scores in domain_scores.items()}
        df_avg = pd.DataFrame(list(avg_scores.items()), columns=["Domain", "Average Score"])

        st.subheader("Average Score per Domain")
        st.bar_chart(df_avg.set_index("Domain"))

        df_meta = pd.DataFrame(metadata)
        if not df_meta.empty:
            st.subheader("Filter by School, Grade, or Subject")
            col1, col2, col3 = st.columns(3)
            school_filter = col1.selectbox("School", ["All"] + sorted(df_meta["School"].dropna().unique().tolist()))
            grade_filter = col2.selectbox("Grade", ["All"] + sorted(df_meta["Grade"].dropna().unique().tolist()))
            subject_filter = col3.selectbox("Subject", ["All"] + sorted(df_meta["Subject"].dropna().unique().tolist()))

            filtered = df_meta.copy()
            if school_filter != "All":
                filtered = filtered[filtered["School"] == school_filter]
            if grade_filter != "All":
                filtered = filtered[filtered["Grade"] == grade_filter]
            if subject_filter != "All":
                filtered = filtered[filtered["Subject"] == subject_filter]

            st.dataframe(filtered)

            st.subheader("Observer Distribution")
            observer_counts = filtered["Observer"].value_counts()
            fig, ax = plt.subplots()
            observer_counts.plot(kind='pie', autopct='%1.1f%%', ax=ax)
            ax.set_ylabel("")
            st.pyplot(fig)
