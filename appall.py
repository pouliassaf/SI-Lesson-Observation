import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime
import os
import statistics
import pandas as pd
import matplotlib.pyplot as plt # This requires matplotlib to be installed
import csv

st.set_page_config(page_title="Lesson Observation Tool", layout="wide")

uploaded_file = None
DEFAULT_FILE = "Teaching Rubric Tool_WeekTemplate.xlsx"
# Check if the default file exists and is readable before trying to open
if os.path.exists(DEFAULT_FILE):
    try:
        uploaded_file = open(DEFAULT_FILE, "rb")
        st.info(f"Using default template workbook: {DEFAULT_FILE}")
    except Exception as e:
        st.error(f"Error opening default template file: {e}")
        uploaded_file = None # Ensure uploaded_file is None if opening fails
else:
    st.warning(f"Default template workbook '{DEFAULT_FILE}' not found. Please upload a workbook.")


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
        try:
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
                     st.error("Error: 'LO 1' template sheet not found in the workbook! Cannot create new sheet.")
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

            lesson_duration = None # Initialize outside try block
            duration_label = "N/A" # Initialize outside try block
            minutes = 0 # Initialize outside try block

            try:
                # Ensure both time_in and time_out are not None before calculating
                if time_in is not None and time_out is not None:
                    # Combine date and time to calculate duration
                    start_time = datetime.combine(datetime.today(), time_in)
                    end_time = datetime.combine(datetime.today(), time_out)

                    # Handle cases where time_out is before time_in (e.g., crossing midnight)
                    if end_time < start_time:
                        end_time += timedelta(days=1) # Assume it's the next day

                    lesson_duration = end_time - start_time
                    minutes = round(lesson_duration.total_seconds() / 60)
                    duration_label = "Full Lesson" if minutes >= 40 else "Walkthrough"
                    st.markdown(f"🕒 **Lesson Duration:** {minutes} minutes — _{duration_label}_")
                else:
                     st.warning("Please enter both 'Time In' and 'Time Out' to calculate duration.")
            except Exception as e:
                st.warning(f"Could not calculate lesson duration: {e}")
                # Import timedelta here if not imported at the top
                from datetime import timedelta


            period = st.selectbox("Period", [f"Period {i}" for i in range(1, 9)])
            obs_type = st.selectbox("Observation Type", ["Individual", "Joint"])

            rubric_domains = {
                "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
            }

            st.markdown("---")
            st.subheader("Rubric Scores")

            # Placeholder for Arabic toggle - not implemented yet
            # arabic_mode = st.toggle("عرض باللغة العربية (Display in Arabic)", False)

            domain_colors = ["#e6f2ff", "#fff2e6", "#e6ffe6", "#f9e6ff", "#ffe6e6", "#f0f0f5", "#e6f9ff", "#f2ffe6", "#ffe6f2"]
            all_element_ratings = {} # To store ratings for feedback generation

            for idx, (domain, (start_cell, count)) in enumerate(rubric_domains.items()):
                background = domain_colors[idx % len(domain_colors)]
                row = int(start_cell[1:])
                domain_title = ws[f'A{row}'].value or domain # Get domain title from A

                st.markdown(f"""
                <div style='background-color:{background};padding:12px;border-radius:10px;margin-bottom:5px;'>
                <h4 style='margin-bottom:5px;'>{domain}: {domain_title}</h4>
                </div>
                """, unsafe_allow_html=True)

                col = start_cell[0]

                for i in range(count):
                    element_row = row + i
                    # Get the element label and description from columns B and C-H
                    label = ws[f"B{element_row}"].value or f"Element {domain[-1]}.{i+1}"

                    # Collect rubric descriptors from columns C to H
                    rubric_descriptors = [ws[f"{chr(ord('C')+j)}{element_row}"].value for j in range(6)]
                    # Filter out None or empty strings and format for display
                    formatted_rubric_text = "\n\n".join([f"**{6-j}:** {desc}" for j, desc in enumerate(rubric_descriptors) if desc])


                    st.markdown(f"**{label}**")
                    with st.expander("Rubric Descriptors"):
                        if formatted_rubric_text:
                            st.markdown(formatted_rubric_text)
                        else:
                            st.info("No rubric descriptors available for this element.")

                    # Get the rating from the user
                    rating = st.selectbox(f"Rating for {label}", [6, 5, 4, 3, 2, 1, "NA"], key=f"{sheet_name}_{domain}_{i}")

                    # Store the rating and description for feedback generation
                    all_element_ratings[f"{domain}_{i}"] = {"label": label, "rating": rating, "description": formatted_rubric_text}

                    # Write the rating to the Excel sheet
                    ws[f"{col}{element_row}"] = rating

            send_feedback = st.checkbox("✉️ Send Feedback to Teacher")

            if st.button("💾 Save Observation"):
                # Ensure essential fields are filled before saving (optional but good practice)
                if not all([observer, teacher, school, grade, subject, students, males, females, time_in, time_out]):
                     st.warning("Please fill in all basic information fields before saving.")
                else: # Proceed only if essential fields are filled
                    ws["Z1"] = "Observer Name"; ws["AA1"] = observer
                    ws["Z2"] = "Teacher"; ws["AA2"] = teacher
                    ws["Z3"] = "Observation Type"; ws["AA3"] = obs_type
                    ws["Z4"] = "Operator"; ws["AA4"] = operator
                    ws["Z5"] = "School"; ws["AA5"] = school
                    ws["Z6"] = "Subject"; ws["AA6"] = subject
                    ws["Z7"] = "Grade"; ws["AA7"] = grade
                    ws["Z8"] = "Gender"; ws["AA8"] = gender
                    ws["Z9"] = "Students"; ws["AA9"] = students # Consider converting to int if needed elsewhere
                    ws["Z10"] = "Males"; ws["AA10"] = males # Consider converting to int if needed elsewhere
                    ws["Z11"] = "Females"; ws["AA11"] = females # Consider converting to int if needed elsewhere
                    ws["Z12"] = "Duration"; ws["AA12"] = duration_label
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
                        # os.remove(save_path) # Consider keeping the file for a bit or saving to a specific temp directory

                    except Exception as e:
                         st.error(f"Error saving workbook: {e}")


                    # Generate and send feedback
                    if send_feedback and teacher_email:
                        feedback_content = f"Dear {teacher},\n\nYour lesson observation from {datetime.now().strftime('%Y-%m-%d')} has been saved.\n\n"
                        feedback_content += f"Observer: {observer}\n"
                        feedback_content += f"Duration: {duration_label}\n"
                        feedback_content += f"Subject: {subject}\n"
                        feedback_content += f"School: {school}\n\n"
                        feedback_content += "Here is a summary of your ratings based on the rubric:\n\n"

                        # Add rubric scores and descriptions to feedback
                        for domain, (start_cell, count) in rubric_domains.items():
                             feedback_content += f"**{domain}: {ws[f'A{int(start_cell[1:])}'].value}**\n"
                             for i in range(count):
                                 element_key = f"{domain}_{i}"
                                 element_info = all_element_ratings.get(element_key)
                                 if element_info:
                                     rating = element_info['rating']
                                     label = element_info['label']
                                     description = element_info['description'] # This is the full formatted rubric text

                                     feedback_content += f"- **{label}:** Rating **{rating}**\n"
                                     # Optional: Include the descriptor for the given rating level
                                     if rating != "NA" and isinstance(rating, int):
                                         try:
                                             # Find the descriptor for the specific rating
                                             # Assuming rubric_descriptors were collected in order 6, 5, 4, 3, 2, 1
                                             descriptor_index = 6 - rating
                                             # Re-read descriptors to get the specific one
                                             element_row = int(start_cell[1:]) + i
                                             specific_descriptor = ws[f"{chr(ord('C')+descriptor_index)}{element_row}"].value
                                             if specific_descriptor:
                                                 feedback_content += f"  *Descriptor for rating {rating}:* {specific_descriptor}\n"
                                         except (IndexError, ValueError):
                                             pass # Handle potential issues with index or rating value

                             feedback_content += "\n" # Add space between domains


                        feedback_content += "Based on these ratings, please review your updated workbook for detailed feedback and areas for development.\n\n"
                        # Placeholder for Support Plan Logic
                        # You would add logic here to check if the teacher's scores
                        # indicate they need a support plan and add a message.
                        # Example: if check_for_support_plan(all_element_ratings):
                        # feedback_content += "Based on your observation, you have been identified for a support plan. Please discuss this with your school leadership.\n\n"


                        feedback_content += "Regards,\nSchool Leadership Team"


                        st.success("Feedback generated (simulated):\n\n" + feedback_content)

                        # Feedback log to sheet
                        try:
                            if "Feedback Log" not in wb.sheetnames:
                                log_ws = wb.create_sheet("Feedback Log")
                                log_ws.append(["Sheet", "Teacher", "Email", "Observer", "School", "Subject", "Date", "Summary"])
                            else:
                                log_ws = wb["Feedback Log"]

                            # Append log entry
                            log_ws.append([
                                sheet_name, teacher, teacher_email, observer, school, subject,
                                datetime.now().strftime("%Y-%m-%d %H:%M"), feedback_content[:200] + ("..." if len(feedback_content) > 200 else "") # Truncate summary
                            ])

                            # Save the workbook again to include the log entry
                            wb.save(save_path)
                            st.success(f"Feedback log updated in {save_path}")

                        except Exception as e:
                            st.error(f"Error updating feedback log in workbook: {e}")

                        # Feedback log as CSV (optional, as it's now in the Excel)
                        # You could generate a CSV from the log sheet if preferred
                        # log_csv_path = "feedback_log.csv"
                        # try:
                        #     with open(log_csv_path, "w", newline="", encoding="utf-8") as f:
                        #         writer = csv.writer(f)
                        #         # Write header from the log sheet
                        #         header = [cell.value for cell in log_ws[1]]
                        #         writer.writerow(header)
                        #         # Write data rows from the log sheet
                        #         for row in log_ws.iter_rows(min_row=2, values_only=True):
                        #             writer.writerow(row)
                        #     with open(log_csv_path, "rb") as f:
                        #         st.download_button("📥 Download Feedback Log (CSV)", f, file_name=log_csv_path)
                        #     # os.remove(log_csv_path) # Clean up the temporary file
                        # except Exception as e:
                        #      st.error(f"Error generating feedback log CSV: {e}")

        except Exception as e:
            st.error(f"Error loading or processing workbook: {e}")


# This elif block is correctly indented at the top level
elif page == "Observation Analytics":
    st.title("Observation Analytics Dashboard")

    if uploaded_file:
        try:
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

                    # Calculate average score for each domain in this sheet and append if ratings exist
                    sheet_domain_avg = {} # Store average for each domain in the current sheet
                    for domain, (start_cell, count) in {
                        "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                        "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
                    }.items():
                        col = start_cell[0]
                        row = int(start_cell[1:])
                        domain_element_ratings = [] # Ratings for elements within THIS domain in THIS sheet
                        for j in range(count):
                            cell_value = ws[f"{col}{row + j}"].value
                            try:
                                # Attempt to convert to float, handles ints and floats
                                rating = float(cell_value)
                                domain_element_ratings.append(rating)
                            except (ValueError, TypeError):
                                # Ignore non-numeric values like "NA" or None
                                pass

                        if domain_element_ratings:
                            # Calculate average for this domain in THIS sheet
                            sheet_domain_avg[domain] = statistics.mean(domain_element_ratings)


                    # Append this sheet's domain averages to the overall list of averages for each domain
                    for domain, avg in sheet_domain_avg.items():
                         domain_scores[domain].append(avg)


                import pandas as pd
                import matplotlib.pyplot as plt

                # Calculate overall averages from domain_scores (which now holds averages per sheet)
                # Only calculate mean if there are scores collected for that domain across sheets
                avg_scores = {domain: round(statistics.mean(scores), 2) if scores else 0 for domain, scores in domain_scores.items() if scores}

                # Ensure all domains are included, using 0 if no scores were collected
                all_domains_list = [f"Domain {i}" for i in range(1, 10)]
                overall_avg_data = [{"Domain": d, "Average Score": avg_scores.get(d, 0)} for d in all_domains_list]
                df_avg = pd.DataFrame(overall_avg_data)

                st.subheader("Average Score per Domain (Across all observations)")
                # Check if there's any data to chart (sum of scores is > 0)
                if not df_avg.empty and df_avg["Average Score"].sum() > 0:
                    st.bar_chart(df_avg.set_index("Domain"))
                else:
                     st.info("No numeric scores found across all observations to calculate overall domain averages.")


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

                    # Filter metadata based on selection
                    filtered_meta_df = df_meta.copy()
                    if school_filter != "All":
                        filtered_meta_df = filtered_meta_df[filtered_meta_df["School"] == school_filter]
                    if grade_filter != "All":
                        filtered_meta_df = filtered_meta_df[filtered_meta_df["Grade"] == grade_filter]
                    if subject_filter != "All":
                        filtered_meta_df = filtered_meta_df[filtered_meta_df["Subject"] == subject_filter]

                    filtered_sheet_names = filtered_meta_df["Sheet"].tolist()

                    # Recalculate domain averages only for the filtered sheets
                    filtered_domain_scores = {domain: [] for domain in all_domains_list}

                    for sheet_name_filtered in filtered_sheet_names:
                         ws_filtered = wb[sheet_name_filtered]
                         sheet_domain_avg_filtered = {}
                         for domain, (start_cell, count) in {
                            "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                            "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
                         }.items():
                            col = start_cell[0]
                            row = int(start_cell[1:])
                            domain_element_ratings = [] # Ratings for elements within THIS domain in THIS sheet
                            for j in range(count):
                                cell_value = ws_filtered[f"{col}{row + j}"].value
                                try:
                                    rating = float(cell_value)
                                    domain_element_ratings.append(rating)
                                except (ValueError, TypeError):
                                    pass # Ignore non-numeric

                            if domain_element_ratings:
                                sheet_domain_avg_filtered[domain] = statistics.mean(domain_element_ratings)

                         # Append this filtered sheet's domain averages
                         for domain, avg in sheet_domain_avg_filtered.items():
                             filtered_domain_scores[domain].append(avg)


                    # Calculate and display filtered averages
                    # Only calculate mean if there are scores collected for that domain after filtering
                    filtered_avg_scores = {domain: round(statistics.mean(scores), 2) if scores else 0 for domain, scores in filtered_domain_scores.items() if scores}

                    # Convert to DataFrame for charting, ensure all domains are present even if avg is 0
                    filtered_avg_data = [{"Domain": d, "Average Score": filtered_avg_scores.get(d, 0)} for d in all_domains_list]
                    df_filtered_avg = pd.DataFrame(filtered_avg_data)


                    st.subheader(f"Average Score per Domain (Filtered)")
                    # Check if there's any data after filtering (sum of scores is > 0)
                    if not df_filtered_avg.empty and df_filtered_avg["Average Score"].sum() > 0:
                        st.bar_chart(df_filtered_avg.set_index("Domain"))
                    else:
                        st.info("No observations matching the selected filters contain numeric scores for domain averages.")


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

        except Exception as e:
             st.error(f"Error loading or processing workbook for analytics: {e}")

    else:
        st.warning("Please upload the workbook or ensure 'Teaching Rubric Tool_WeekTemplate.xlsx' exists to view analytics.")

