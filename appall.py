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
            # Ensure the template sheet "LO 1" exists before copying
            if "LO 1" in wb.sheetnames:
                 wb.copy_worksheet(wb["LO 1"]).title = sheet_name
                 st.success(f"Created new sheet: {sheet_name}")
            else:
                 st.error("Error: 'LO 1' template sheet not found in the workbook!")
                 st.stop() # Stop execution if template is missing

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
                "Al Qeyam Charter School", "Al Qeyam Charter School", "Al Nayfa Charter Kindergarten", "Al Salam Charter School",
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

        lesson_duration = None # Initialize outside try block
        duration_label = "N/A" # Initialize outside try block
        minutes = 0 # Initialize outside try block

        try:
            # Ensure both time_in and time_out are not None before calculating
            if time_in is not None and time_out is not None:
                lesson_duration = datetime.combine(datetime.today(), time_out) - datetime.combine(datetime.today(), time_in)
                minutes = round(lesson_duration.total_seconds() / 60)
                duration_label = "Full Lesson" if minutes >= 40 else "Walkthrough"
                st.markdown(f"🕒 **Lesson Duration:** {minutes} minutes — _{duration_label}_")
            else:
                 st.warning("Please enter both 'Time In' and 'Time Out' to calculate duration.")
        except Exception as e:
            st.warning(f"Could not calculate lesson duration: {e}")


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
            # Safely get the domain title from column A
            domain_title = ws[f'A{row}'].value or domain # Use domain name as fallback
            st.markdown(f"**{domain_title}**")
            for i in range(count):
                label = ws[f"B{row + i}"].value or f"Element {domain[-1]}.{i+1}"
                rating = st.selectbox(f"Rating for {label}", [6, 5, 4, 3, 2, 1, "NA"], key=f"{sheet_name}_{domain}_{i}") # Added sheet_name to key for uniqueness across sheets
                ws[f"{col}{row + i}"] = rating

        send_feedback = st.checkbox("✉️ Send Feedback to Teacher")

        if st.button("💾 Save Observation"):
            # Ensure essential fields are filled before saving (optional but good practice)
            if not all([observer, teacher, school, grade, subject, students, males, females, time_in, time_out]):
                 st.warning("Please fill in all basic information fields before saving.")
                 #st.stop() # Don't stop, just warn. Let them fill and click save again.

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
            ws["AA9"] = students # Consider converting to int if needed elsewhere
            ws["Z10"] = "Males"
            ws["AA10"] = males # Consider converting to int if needed elsewhere
            ws["Z11"] = "Females"
            ws["AA11"] = females # Consider converting to int if needed elsewhere
            ws["Z12"] = "Duration"
            ws["AA12"] = duration_label
            ws["Z13"] = "Time In"
            # Check if time_in is not None before formatting
            ws["AA13"] = time_in.strftime("%H:%M") if time_in else "N/A"
            ws["Z14"] = "Time Out"
             # Check if time_out is not None before formatting
            ws["AA14"] = time_out.strftime("%H:%M") if time_out else "N/A"


            save_path = f"updated_{sheet_name}.xlsx"
            try:
                wb.save(save_path)
                st.success(f"Observation data saved to {sheet_name} in {save_path}")
                with open(save_path, "rb") as f:
                    st.download_button("📥 Download updated workbook", f, file_name=save_path)
                os.remove(save_path) # Clean up the temporary file
            except Exception as e:
                 st.error(f"Error saving workbook: {e}")


            # **Corrected feedback string formatting**
            if send_feedback and teacher_email:
                feedback = (
                    f"Dear {teacher},\n\n"  # Use \n\n for blank lines
                    "Your lesson observation has been saved.\n" # Use \n for single lines
                    f"Observer: {observer}\n"
                    f"Duration: {duration_label}\n"
                    f"Subject: {subject}\n"
                    f"School: {school}\n\n" # Use \n\n
                    "Based on rubric ratings, please review your updated workbook for details.\n\n" # Use \n\n
                    "Regards,\n" # Use \n
                    "School Leadership Team"
                )
                # This line's indentation was already corrected
                st.success("Feedback generated (simulated):\n\n" + feedback)


# This elif block is correctly indented at the top level
elif page == "Observation Analytics":
    st.title("Observation Analytics Dashboard")

    if uploaded_file:
        # Use data_only=True to get calculated values from the Excel file
        wb = load_workbook(uploaded_file, data_only=True)
        sheets = [s for s in wb.sheetnames if s.startswith("LO ")]

        if not sheets:
            st.warning("No 'LO ' sheets found in the workbook for analytics.")
        else:
            domain_scores = {domain: [] for domain in [f"Domain {i}" for i in range(1, 10)]}
            metadata = []

            for sheet in sheets:
                ws = wb[sheet]
                row_info = {
                    "Sheet": sheet, # Add sheet name for reference
                    "Observer": ws["AA1"].value,
                    "Teacher": ws["AA2"].value, # Added Teacher name
                    "Observation Type": ws["AA3"].value, # Added Observation Type
                    "Operator": ws["AA4"].value, # Added Operator
                    "School": ws["AA5"].value,
                    "Subject": ws["AA6"].value,
                    "Grade": ws["AA7"].value,
                    "Gender": ws["AA8"].value, # Added Gender
                    "Students": ws["AA9"].value, # Added Student counts
                    "Males": ws["AA10"].value,
                    "Females": ws["AA11"].value,
                    "Duration": ws["AA12"].value, # Added duration label
                    "Time In": ws["AA13"].value, # Added times
                    "Time Out": ws["AA14"].value,
                 }
                metadata.append(row_info)

                # Collect individual rubric element scores (convert to float if possible, ignore non-numeric)
                sheet_domain_ratings = {domain: [] for domain in domain_scores.keys()}
                for domain, (start_cell, count) in {
                    "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                    "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
                }.items():
                    col = start_cell[0]
                    row = int(start_cell[1:])
                    for j in range(count):
                        cell_value = ws[f"{col}{row + j}"].value
                        # Try converting to float, if it fails or is None, skip it
                        try:
                            rating = float(cell_value)
                            sheet_domain_ratings[domain].append(rating)
                        except (ValueError, TypeError):
                            # Ignore non-numeric values like "NA"
                            pass

                # Calculate average score for each domain in this sheet and append if ratings exist
                for domain, ratings in sheet_domain_ratings.items():
                    if ratings:
                        domain_scores[domain].append(statistics.mean(ratings))


            import pandas as pd
            import matplotlib.pyplot as plt

            # Ensure dataframe is created even if no scores were numeric
            avg_scores = {domain: round(statistics.mean(scores), 2) if scores else 0 for domain, scores in domain_scores.items()}
            df_avg = pd.DataFrame(list(avg_scores.items()), columns=["Domain", "Average Score"])

            st.subheader("Average Score per Domain (Across all observations)")
            st.bar_chart(df_avg.set_index("Domain"))

            df_meta = pd.DataFrame(metadata)
            if not df_meta.empty:
                st.subheader("Observation Data Summary")
                st.dataframe(df_meta) # Show the full metadata table

                st.subheader("Filter and Analyze")
                # Use unique values from the dataframe for filters
                col1, col2, col3 = st.columns(3)
                school_filter = col1.selectbox("Filter by School", ["All"] + sorted(df_meta["School"].dropna().unique().tolist()))
                grade_filter = col2.selectbox("Filter by Grade", ["All"] + sorted(df_meta["Grade"].dropna().unique().tolist()))
                subject_filter = col3.selectbox("Filter by Subject", ["All"] + sorted(df_meta["Subject"].dropna().unique().tolist()))

                # Re-calculate filtered averages based on metadata filter
                filtered_meta_df = df_meta.copy()
                if school_filter != "All":
                    filtered_meta_df = filtered_meta_df[filtered_meta_df["School"] == school_filter]
                if grade_filter != "All":
                    filtered_meta_df = filtered_meta_df[filtered_meta_df["Grade"] == grade_filter]
                if subject_filter != "All":
                    filtered_meta_df = filtered_meta_df[filtered_meta_df["Subject"] == subject_filter]

                filtered_sheet_names = filtered_meta_df["Sheet"].tolist()

                filtered_domain_scores = {domain: [] for domain in domain_scores.keys()}

                # Re-process sheets that match the filter criteria
                for sheet_name_filtered in filtered_sheet_names:
                     # Need to re-load the workbook or access sheets from the original loaded wb
                     # Accessing from wb (loaded data_only=True) is fine
                     ws_filtered = wb[sheet_name_filtered]
                     for domain, (start_cell, count) in {
                        "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                        "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
                     }.items():
                        col = start_cell[0]
                        row = int(start_cell[1:])
                        ratings_filtered = []
                        for j in range(count):
                            cell_value = ws_filtered[f"{col}{row + j}"].value
                            try:
                                rating = float(cell_value)
                                ratings_filtered.append(rating)
                            except (ValueError, TypeError):
                                pass # Ignore non-numeric

                        if ratings_filtered:
                            filtered_domain_scores[domain].append(statistics.mean(ratings_filtered))

                # Calculate and display filtered averages
                filtered_avg_scores = {domain: round(statistics.mean(scores), 2) if scores else 0 for domain, scores in filtered_domain_scores.items()}
                df_filtered_avg = pd.DataFrame(list(filtered_avg_scores.items()), columns=["Domain", "Average Score"])

                st.subheader(f"Average Score per Domain (Filtered)")
                if not df_filtered_avg.empty and df_filtered_avg["Average Score"].sum() > 0: # Check if there's any data after filtering
                    st.bar_chart(df_filtered_avg.set_index("Domain"))
                else:
                    st.info("No observations match the selected filters with numeric domain scores.")


                st.subheader("Observer Distribution (Filtered)")
                # Use the already filtered_meta_df for observer counts
                if not filtered_meta_df.empty:
                    observer_counts = filtered_meta_df["Observer"].value_counts()
                    if not observer_counts.empty:
                         fig, ax = plt.subplots()
                         observer_counts.plot(kind='pie', autopct='%1.1f%%', ax=ax)
                         ax.set_ylabel("") # Hide the default y-label for pie charts
                         st.pyplot(fig)
                    else:
                        st.info("No observer data found for the selected filters.")
                else:
                    st.info("No observation data found for the selected filters.")


    else:
        st.warning("Please upload the workbook or ensure 'Teaching Rubric Tool_WeekTemplate.xlsx' exists to view analytics.")
