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
allowed_domains = ["@charterschools.ae", "@adek.gov.ae", "@nceducation.com"]

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
                st.download_button("📄 Download School Trends Report (CSV)", data=trend_io.getvalue(), file_name="school_trends.csv", mime="text/csv")

                st.subheader("📤 Export Filtered Data")

                st.subheader("🧾 Reflection Summary Report")
                st.markdown("Shows strengths and areas for improvement per group.")

                with st.expander("Generate by School"):
                    school_reflection = filtered_df.groupby("School")["Overall Avg"].mean().sort_values(ascending=False)
                    st.write("**Areas of Strength**")
                    st.write(school_reflection.head(3))
                    st.write("**Areas for Improvement**")
                    st.write(school_reflection.tail(3))

                if "Subject" in filtered_df.columns:
                    with st.expander("Generate by Subject"):
                        subject_reflection = filtered_df.groupby("Subject")["Overall Avg"].mean().sort_values(ascending=False)
                        st.write("**Areas of Strength**")
                        st.write(subject_reflection.head(3))
                        st.write("**Areas for Improvement**")
                        st.write(subject_reflection.tail(3))

                if "Grade" in filtered_df.columns:
                    with st.expander("Generate by Grade"):
                        grade_reflection = filtered_df.groupby("Grade")["Overall Avg"].mean().sort_values(ascending=False)
                        st.write("**Areas of Strength**")
                        st.write(grade_reflection.head(3))
                        st.write("**Areas for Improvement**")
                        st.write(grade_reflection.tail(3))

                
                import io
                reflection_buffer = io.StringIO()
import io
reflection_buffer = io.StringIO()
reflection_buffer.write("Reflection Summary Report\n\n")

if not school_reflection.empty:
    reflection_buffer.write("School - Areas of Strength:\n")
    reflection_buffer.write(school_reflection.head(3).to_string())
    reflection_buffer.write("\n\nSchool - Areas for Improvement:\n")
    reflection_buffer.write(school_reflection.tail(3).to_string())
    reflection_buffer.write("\n\n")

if 'subject_reflection' in locals() and not subject_reflection.empty:
    reflection_buffer.write("Subject - Areas of Strength:\n")
    reflection_buffer.write(subject_reflection.head(3).to_string())
    reflection_buffer.write("\n\nSubject - Areas for Improvement:\n")
    reflection_buffer.write(subject_reflection.tail(3).to_string())
    reflection_buffer.write("\n\n")

if 'grade_reflection' in locals() and not grade_reflection.empty:
    reflection_buffer.write("Grade - Areas of Strength:\n")
    reflection_buffer.write(grade_reflection.head(3).to_string())
    reflection_buffer.write("\n\nGrade - Areas for Improvement:\n")
    reflection_buffer.write(grade_reflection.tail(3).to_string())
    reflection_buffer.write("\n\n")

from fpdf import FPDF

pdf = FPDF()

if os.path.exists("logos"):
    logo_path = (
        f"logos/{school}.png" if os.path.exists(f"logos/{school}.png") else
        f"logos/{school}.jpg" if os.path.exists(f"logos/{school}.jpg") else
        f"logos/{school}.jpeg" if os.path.exists(f"logos/{school}.jpeg") else
        "logos/default.png" if os.path.exists("logos/default.png") else
        "logos/default.jpg" if os.path.exists("logos/default.jpg") else
        "logos/default.jpeg"
    )
    if os.path.exists(logo_path):
        pdf.image(logo_path, x=170, y=8, w=30)

pdf.add_page()
pdf.set_font("Arial", size=12)
pdf.cell(200, 10, txt="Lesson Observation Summary", ln=True, align='C')
pdf.ln(10)
pdf.multi_cell(0, 10, txt=(
    f"Observer: {observer}\n"
    f"Teacher: {teacher}\n"
    f"Subject: {subject}\n"
    f"School: {school}\n"
    f"Grade: {grade}\n"
    f"Gender: {gender}\n"
    f"Date: {date}\n"
    f"Duration: {duration_label}\n"
    f"Period: {period}\n"
    f"Observation Type: {obs_type}"
))
pdf.ln(5)
pdf.cell(200, 10, txt=f"Overall Average: {overall_avg if all_scores else 'N/A'}", ln=True)
pdf.cell(200, 10, txt=f"Final Judgment: {overall_judgment if all_scores else 'N/A'}", ln=True)
pdf.ln(5)
pdf.multi_cell(0, 10, txt=(f"General Notes:\n{overall_notes}"))
))
pdf.ln(5)
pdf.multi_cell(0, 10, txt=support_plan)
pdf.set_y(-20)
pdf.set_font("Arial", size=8)
pdf.cell(0, 10, txt=f"{school} • {date.strftime('%Y-%m-%d')}", ln=True, align='C')

pdf_output = pdf.output(dest='S').encode('latin-1')

        pdf_lang = st.radio("Select PDF language", ["English", "Arabic"], horizontal=True)

        # PDF already initialized and logo added above
        # pdf = FPDF()
        # pdf.add_page()
        pdf.set_font("Arial", size=12)
        if pdf_lang == "Arabic":
            pdf.cell(200, 10, txt="ملخص ملاحظة الحصة الدراسية", ln=True, align='C')
            pdf.ln(10)
            pdf.set_font("Arial", size=12)
pdf.multi_cell(0, 10, txt=(f"ملاحظات عامة:\n{overall_notes}"))

))
            pdf.ln(5)
            support_plan = "الخطوات التالية:
"
            if overall_judgment in ["Weak", "Very Weak"]:
                support_plan += "- يجب تحديد جلسة متابعة خلال أسبوعين.
- يُنصح بالتطوير المهني المستهدف.
- تقديم دعم صفّي وفرص للملاحظة الزميلية."
            elif overall_judgment == "Acceptable":
                support_plan += "- التشجيع على التفكير الذاتي في مجالات التحسين.
- الانضمام إلى مجتمعات التعلم المهني.
- تتبع ملاحظات المتابعة خلال الفصل."
            elif overall_judgment in ["Good", "Very Good"]:
                support_plan += "- الحفاظ على الاتساق في الأداء.
- دعم الزملاء من خلال الإرشاد."
            elif overall_judgment == "Outstanding":
                support_plan += "- قيادة ورش العمل المهنية.
- مشاركة أفضل الممارسات.
- الإسهام في مشاريع التحسين."
            else:
                support_plan += "- لا يوجد حكم نهائي. يرجى مراجعة المدخلات."
        else:
            pdf.cell(200, 10, txt="Lesson Observation Summary", ln=True, align='C')
            pdf.ln(10)
            pdf.multi_cell(0, 10, txt=(
    f"Observer: {observer}
"
    f"Teacher: {teacher}
"
    f"Subject: {subject}
"
    f"School: {school}
"
    f"Grade: {grade}
"
    f"Gender: {gender}
"
    f"Date: {date}
"
    f"Duration: {duration_label}
"
    f"Period: {period}
"
    f"Observation Type: {obs_type}"
))
            pdf.ln(5)
            pdf.cell(200, 10, txt=f"Overall Average: {overall_avg if all_scores else 'N/A'}", ln=True)
            pdf.cell(200, 10, txt=f"Final Judgment: {overall_judgment if all_scores else 'N/A'}", ln=True)
            pdf.ln(5)
            pdf.multi_cell(0, 10, txt=f"General Notes:
{overall_notes}")
            pdf.ln(5)
            support_plan = "Next Steps:
"
            if overall_judgment in ["Weak", "Very Weak"]:
                support_plan += "- A follow-up coaching session should be scheduled within 2 weeks.
- Targeted professional development should be prioritized.
- Provide classroom support and peer observation opportunities."
            elif overall_judgment == "Acceptable":
                support_plan += "- Encourage reflection on areas of improvement.
- Recommend joining PLC sessions.
- Track follow-up observations within the term."
            elif overall_judgment in ["Good", "Very Good"]:
                support_plan += "- Maintain consistency across lessons.
- Support other teachers through mentorship or peer reviews."
            elif overall_judgment == "Outstanding":
                support_plan += "- Consider leading PD sessions.
- Share exemplary practices across teams.
- Contribute to strategic improvement projects."
            else:
                support_plan += "- No judgment available. Please review rubric input."

        pdf.multi_cell(0, 10, txt=support_plan)
        pdf_output = pdf.output(dest='S').encode('latin-1')

        st.download_button(
            label="📥 Download Form Summary (PDF)",
            data=pdf_output,
            file_name=f"{sheet_name}_summary.pdf",
            mime="application/pdf"
        )

        if st.button("📤 Email PDF to Teacher"):
            pdf_email_msg = MIMEMultipart()
            pdf_email_msg['From'] = sender_email
            pdf_email_msg['To'] = teacher_email
            pdf_email_msg['Cc'] = sender_email
            pdf_email_msg['Subject'] = f"Lesson Observation PDF Summary – {teacher}"
            pdf_email_msg.attach(MIMEText("Attached is your observation summary in PDF format.", 'plain'))
            from email.mime.application import MIMEApplication
            attachment = MIMEApplication(pdf_output, _subtype="pdf")
            attachment.add_header('Content-Disposition', 'attachment', filename=f"{sheet_name}_summary.pdf")
            pdf_email_msg.attach(attachment)

            try:
                smtp = smtplib.SMTP(smtp_server, smtp_port)
                smtp.starttls()
                smtp.login(smtp_user, smtp_pass)
                smtp.sendmail(sender_email, [teacher_email, sender_email], pdf_email_msg.as_string())
                smtp.quit()
                st.success("PDF sent to teacher and observer successfully!")

                # Log email send
                if "Email Log" not in wb.sheetnames:
                    email_log_ws = wb.create_sheet("Email Log")
                    email_log_ws.append(["Timestamp", "Observer", "Observer Email", "Teacher", "Teacher Email", "Type", "Status"])
                else:
                    email_log_ws: Worksheet = wb["Email Log"]

                email_log_ws.append([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    observer,
                    sender_email,
                    teacher,
                    teacher_email,
                    "PDF sent",
                    "Success"
                ])

                wb.save(DEFAULT_FILE)
            except Exception as e:
                st.error(f"PDF email failed: {e}")
            mime="application/pdf"
        )

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








































  

  










