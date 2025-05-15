import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime
import os
import statistics




uploaded_file = None
DEFAULT_FILE = "Teaching Rubric Tool_WeekTemplate.xlsx"
if not uploaded_file and os.path.exists(DEFAULT_FILE):
    uploaded_file = open(DEFAULT_FILE, "rb")
    st.info("Using default template workbook.")

st.set_page_config(page_title="Lesson Observation Tool", layout="wide")

page = st.sidebar.selectbox("Choose a page:", ["Lesson Input", "Observation Analytics"])

if page == "Lesson Input":
    st.title("Weekly Lesson Observation Input Tool")

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

        st.session_state.update({
            "observer": observer, "teacher": teacher, "operator": operator, "school": school,
            "grade": grade, "subject": subject, "gender": gender, "students": students,
            "males": males, "females": females, "time_in": time_in, "time_out": time_out,
            "duration_label": duration_label, "period": period, "obs_type": obs_type
        })

    

    





elif page == "Observation Analytics":
    import pandas as pd
    import matplotlib.pyplot as plt

    st.title("Observation Analytics")

    if "Email Log" in wb.sheetnames:
        st.subheader("📬 Email Report Summary")
        email_ws = wb["Email Log"]
        email_data = list(email_ws.values)
        email_headers, *email_rows = email_data
        email_df = pd.DataFrame(email_rows, columns=email_headers)

        st.markdown("**Total Feedback Reports Sent:**")
        st.metric("Total Sent", len(email_df))

        email_by_teacher = email_df["Teacher"].value_counts()
        st.markdown("**Most Frequently Contacted Teachers:**")
        st.bar_chart(email_by_teacher.head(10))

        email_by_observer = email_df["Observer"].value_counts()
        st.markdown("**Emails Sent per Observer:**")
        st.bar_chart(email_by_observer)

    if os.path.exists(DEFAULT_FILE):
        wb = load_workbook(DEFAULT_FILE)
        if "Observation Log" in wb.sheetnames:
            ws = wb["Observation Log"]
            data = list(ws.values)
            headers, *rows = data
            df = pd.DataFrame(rows, columns=headers)
            df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors='coerce')
            st.success("Observation Log loaded successfully.")

            if not df.empty:
                st.subheader("Filter Observations")
                with st.expander("🔎 Filter Options"):
                    selected_school = st.multiselect("Filter by School", sorted(df["School"].dropna().unique()))
                    selected_subject = st.multiselect("Filter by Subject", sorted(df["Subject"].dropna().unique()) if "Subject" in df.columns else [])
                    selected_grade = st.multiselect("Filter by Grade", sorted(df["Grade"].dropna().unique()) if "Grade" in df.columns else [])

                    filtered_df = df.copy()
                    if selected_school:
                        filtered_df = filtered_df[filtered_df["School"].isin(selected_school)]
                    if selected_subject:
                        filtered_df = filtered_df[filtered_df["Subject"].isin(selected_subject)]
                    if selected_grade:
                        filtered_df = filtered_df[filtered_df["Grade"].isin(selected_grade)]

                st.subheader("Observer Summary")
                observer_avg = df.groupby("Observer")["Overall Avg"].mean().sort_values(ascending=False)
                st.bar_chart(observer_avg)

                observer_count = df["Observer"].value_counts()
                st.markdown("#### Observations per Observer")
                st.bar_chart(observer_count)

                st.subheader("Overall Judgment Distribution")
                st.bar_chart(filtered_df["Final Judgment"].value_counts())

                st.subheader("Average Score by School")
                avg_school = filtered_df.groupby("School")["Overall Avg"].mean().sort_values()
                st.bar_chart(avg_school)

                st.subheader("Average Score by Subject")
                if "Subject" in filtered_df.columns:
                    avg_subject = filtered_df.groupby("Subject")["Overall Avg"].mean().sort_values()
                    st.bar_chart(avg_subject)

                st.subheader("Average Score by Grade")
                if "Grade" in filtered_df.columns:
                    avg_grade = filtered_df.groupby("Grade")["Overall Avg"].mean().sort_values()
                    st.bar_chart(avg_grade)

                st.subheader("📈 School Comparison Over Time")
                df_time = df.dropna(subset=["Timestamp"])
                df_time['Month'] = df_time['Timestamp'].dt.to_period('M').astype(str)
                trend_df = df_time.groupby(['Month', 'School'])['Overall Avg'].mean().reset_index()
                for school_name in trend_df['School'].unique():
                    st.markdown(f"#### {school_name}")
                    school_trend = trend_df[trend_df['School'] == school_name]
                    fig, ax = plt.subplots()
                    ax.plot(school_trend['Month'], school_trend['Overall Avg'], marker='o')
                    ax.set_title(f"Average Observation Score Over Time for {school_name}")
                    ax.set_xlabel("Month")
                    ax.set_ylabel("Average Score")
                    ax.tick_params(axis='x', rotation=45)
                    st.pyplot(fig)

                trend_csv = trend_df.pivot(index='Month', columns='School', values='Overall Avg').fillna('').reset_index()
                import io
                trend_io = io.StringIO()
                trend_csv.to_csv(trend_io, index=False)

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
