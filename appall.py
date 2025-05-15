import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime
import os
import statistics
import pandas as pd
import matplotlib.pyplot as plt
import io
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

st.set_page_config(page_title="Lesson Observation Tool", layout="wide")

page = st.sidebar.selectbox("Choose a page:", ["Lesson Input", "Observation Analytics"])

if page == "Lesson Input":
    st.title("Weekly Lesson Observation Input Tool")

email = st.text_input("Enter your school email to continue")
allowed_domains = ["@charterschools.ae", "@adek.gov.ae", "@nceducation.com"]

if not any(email.endswith(domain) for domain in allowed_domains):
    st.warning("Access restricted. Please use an authorized school email.")
    st.stop()

DEFAULT_FILE = "Teaching Rubric Tool_WeekTemplate.xlsx"
if os.path.exists(DEFAULT_FILE):
    uploaded_file = open(DEFAULT_FILE, "rb")
    st.info("Using default template workbook.")
    wb = load_workbook(DEFAULT_FILE)

if page == "Observation Analytics":
    st.title("Observation Analytics")

    if "Email Log" in wb.sheetnames:
        st.subheader("📬 Email Report Summary")
        email_ws = wb["Email Log"]
        email_data = list(email_ws.values)
        email_headers, *email_rows = email_data
        email_df = pd.DataFrame(email_rows, columns=email_headers)

        st.metric("Total Sent", len(email_df))
        st.bar_chart(email_df["Teacher"].value_counts().head(10))
        st.bar_chart(email_df["Observer"].value_counts())

    if "Observation Log" in wb.sheetnames:
        ws = wb["Observation Log"]
        data = list(ws.values)
        headers, *rows = data
        df = pd.DataFrame(rows, columns=headers)
        df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors='coerce')
        st.success("Observation Log loaded successfully.")

        if not df.empty:
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
            st.bar_chart(df.groupby("Observer")["Overall Avg"].mean().sort_values(ascending=False))
            st.bar_chart(df["Observer"].value_counts())

            st.subheader("Overall Judgment Distribution")
            st.bar_chart(filtered_df["Final Judgment"].value_counts())

            st.subheader("Average Score by School")
            st.bar_chart(filtered_df.groupby("School")["Overall Avg"].mean().sort_values())

            if "Subject" in filtered_df.columns:
                st.subheader("Average Score by Subject")
                st.bar_chart(filtered_df.groupby("Subject")["Overall Avg"].mean().sort_values())

            if "Grade" in filtered_df.columns:
                st.subheader("Average Score by Grade")
                st.bar_chart(filtered_df.groupby("Grade")["Overall Avg"].mean().sort_values())

            df_time = df.dropna(subset=["Timestamp"])
            df_time['Month'] = df_time['Timestamp'].dt.to_period('M').astype(str)
            trend_df = df_time.groupby(['Month', 'School'])["Overall Avg"].mean().reset_index()
            for school_name in trend_df['School'].unique():
                school_trend = trend_df[trend_df['School'] == school_name]
                fig, ax = plt.subplots()
                ax.plot(school_trend['Month'], school_trend['Overall Avg'], marker='o')
                ax.set_title(f"Average Observation Score Over Time for {school_name}")
                ax.set_xlabel("Month")
                ax.set_ylabel("Average Score")
                ax.tick_params(axis='x', rotation=45)
                st.pyplot(fig)

            trend_csv = trend_df.pivot(index='Month', columns='School', values='Overall Avg').fillna('').reset_index()
            trend_io = io.StringIO()
            trend_csv.to_csv(trend_io, index=False)
            st.download_button("📄 Download School Trends Report (CSV)", data=trend_io.getvalue(), file_name="school_trends.csv", mime="text/csv")

            st.subheader("📤 Export Filtered Data")
            school_reflection = filtered_df.groupby("School")["Overall Avg"].mean().sort_values(ascending=False)
            subject_reflection = filtered_df.groupby("Subject")["Overall Avg"].mean().sort_values(ascending=False) if "Subject" in filtered_df.columns else pd.Series()
            grade_reflection = filtered_df.groupby("Grade")["Overall Avg"].mean().sort_values(ascending=False) if "Grade" in filtered_df.columns else pd.Series()

            pdf_lang = st.radio("Select PDF language", ["English", "Arabic"], horizontal=True)
            teacher = st.text_input("Teacher Name for PDF")
            teacher_email = st.text_input("Teacher Email")
            observer = st.text_input("Observer Name")
            sender_email = email
            smtp_server = "smtp.office365.com"
            smtp_port = 587
            smtp_user = sender_email
            smtp_pass = st.text_input("Enter your email password to send the PDF", type="password")

            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            if pdf_lang == "Arabic":
                pdf.multi_cell(0, 10, txt="ملخص ملاحظة الحصة الدراسية")
                pdf.multi_cell(0, 10, txt="ملاحظات عامة:\n" + str("Sample notes here"))
                support_plan = "الخطوات التالية:\n"
                support_plan += "- يُنصح بالتطوير المهني المستهدف.\n"
                support_plan += "- تقديم دعم صفّي وفرص للملاحظة الزميلية."
            else:
                pdf.multi_cell(0, 10, txt="Lesson Observation Summary")
                pdf.multi_cell(0, 10, txt="General Notes:\n" + str("Sample notes here"))
                support_plan = "Next Steps:\n"
                support_plan += "- Targeted professional development should be prioritized.\n"
                support_plan += "- Provide classroom support and peer observation opportunities."

            pdf.multi_cell(0, 10, txt=support_plan)
            pdf_output = pdf.output(dest='S').encode('latin-1')

            st.download_button("📥 Download Form Summary (PDF)", data=pdf_output, file_name="summary.pdf", mime="application/pdf")

            if st.button("📤 Email PDF to Teacher"):
                try:
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = teacher_email
                    msg['Cc'] = sender_email
                    msg['Subject'] = f"Lesson Observation PDF Summary – {teacher}"
                    msg.attach(MIMEText("Attached is your observation summary in PDF format.", 'plain'))
                    attachment = MIMEApplication(pdf_output, _subtype="pdf")
                    attachment.add_header('Content-Disposition', 'attachment', filename=f"{teacher}_summary.pdf")
                    msg.attach(attachment)

                    smtp = smtplib.SMTP(smtp_server, smtp_port)
                    smtp.starttls()
                    smtp.login(smtp_user, smtp_pass)
                    smtp.sendmail(sender_email, [teacher_email, sender_email], msg.as_string())
                    smtp.quit()
                    st.success("PDF sent to teacher and observer successfully!")
                except Exception as e:
                    st.error(f"PDF email failed: {e}")
