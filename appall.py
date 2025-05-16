import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime, timedelta, date
import os
import statistics
import pandas as pd
import matplotlib.pyplot as plt
import csv
import math
import io

# Import ReportLab modules for PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image # Import Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors

# --- Set Streamlit Page Config (MUST BE THE FIRST STREAMLIT COMMAND) ---
st.set_page_config(page_title="Lesson Observation Tool", layout="wide")

# --- Logo File Paths ---
# Define a dictionary mapping school names to logo file paths
# Ensure these paths are correct relative to your script location
# Add more school logos here as needed
LOGO_PATHS = {
    "Default": "logos/default.jpeg", # Default logo for schools not listed
    "Al Bayan School": "logos/CS_Al Bayan Charter School_Logo.png",
    "Al Bayraq School": "logos/CS_Al Bayraq_Logo_PNG.png",
    "Al Dhaher School": "logos/CS_Al Dhaher_Logo_PNG.png",
    "Al Hosoon School": "logos/CS_Al Hosoon Charter School_Logo.png",
    "Al Mutanabi School": "logos/CS_Al Mutanabi_Logo_PNG.png",
    "Al Nahdha School": "logos/CS_Al Nahdha_Logo_PNG.png",
    "Jern Yafoor School": "logos/CS_Jern Yafoor_Logo_PNG.png",
    "Maryam Bint Omran School": "logos/CS_Maryam Bint Omran_Logo_PNG.png",
    # Add other school logos here following the pattern:
    # "School Name as it appears in the app": "logos/your_logo_file.png",
}


# --- Text Strings for Localization ---
# You need to replace the placeholder Arabic strings with actual translations
en_strings = {
    "page_title": "Lesson Observation Tool", # Reverted to Lesson Observation
    "sidebar_select_page": "Choose a page:",
    "page_lesson_input": "Lesson Observation Input", # Reverted to Lesson Observation
    "page_analytics": "Lesson Observation Analytics", # Reverted to Lesson Observation
    "title_lesson_input": "Weekly Lesson Observation Input Tool", # Reverted to Lesson Observation
    "info_default_workbook": "Using default template workbook:",
    "warning_default_not_found": "Default template workbook '{}' not found. Please upload a workbook.",
    "error_opening_default": "Error opening default template file:",
    "success_lo_sheets_found": "Found {} LO sheets in workbook.",
    "select_sheet_or_create": "Select existing LO sheet or create a new one:",
    "option_create_new": "Create new",
    "success_sheet_created": "Created new sheet: {}",
    "error_template_not_found": "Error: 'LO 1' template sheet not found in the workbook! Cannot create new sheet.",
    "subheader_filling_data": "Filling data for: {}",
    "label_observer_name": "Observer Name",
    "label_teacher_name": "Teacher Name",
    "label_teacher_email": "Teacher Email",
    "label_operator": "Operator",
    "label_school_name": "School Name",
    "label_grade": "Grade",
    "label_subject": "Subject",
    "label_gender": "Gender",
    "label_students": "Number of Students",
    "label_males": "Number of Males",
    "label_females": "Number of Females",
    "label_time_in": "Time In",
    "label_time_out": "Time Out",
    "label_lesson_duration": "🕒 **Lesson Duration:** {} minutes — _{}_",
    "duration_full_lesson": "Full Lesson",
    "duration_walkthrough": "Walkthrough",
    "warning_calculate_duration": "Please enter both 'Time In' and 'Time Out' to calculate duration.",
    "warning_could_not_calculate_duration": "Could not calculate lesson duration:",
    "label_period": "Period",
    "label_obs_type": "Observation Type", # Reverted to Observation Type
    "option_individual": "Individual",
    "option_joint": "Joint",
    "subheader_rubric_scores": "Rubric Scores",
    "expander_rubric_descriptors": "Rubric Descriptors",
    "info_no_descriptors": "No rubric descriptors available for this element.",
    "label_rating_for": "Rating for {}",
    "checkbox_send_feedback": "✉️ Send Feedback to Teacher",
    "button_save_observation": "💾 Save Observation", # Reverted to Save Observation
    "warning_fill_essential": "Please fill in all basic information fields before saving.",
    "success_data_saved": "Observation data saved to {} in {}", # Reverted to Observation data
    "error_saving_workbook": "Error saving workbook:",
    "download_workbook": "📥 Download updated workbook",
    "feedback_subject": "Lesson Observation Feedback", # Reverted to Lesson Observation
    "feedback_greeting": "Dear {},\n\nYour lesson observation from {} has been saved.\n\n", # Reverted to lesson observation
    "feedback_observer": "Observer: {}\n",
    "feedback_duration": "Duration: {}\n",
    "feedback_subject_fb": "Subject: {}\n", # Renamed to avoid conflict with label_subject
    "feedback_school": "School: {}\n\n",
    "feedback_summary_header": "Here is a summary of your ratings based on the rubric:\n\n",
    "feedback_domain_header": "**{}: {}**\n", # Domain number and title
    "feedback_element_rating": "- **{}:** Rating **{}**\n", # Element label and rating
    "feedback_descriptor_for_rating": "  *Descriptor for rating {}:* {}\n", # Descriptor for specific rating
    "feedback_overall_score": "\n**Overall Average Score:** {:.2f}\n\n", # Added for overall score
    "feedback_domain_average": "  *Domain Average:* {:.2f}\n", # Added for domain average
    "feedback_performance_summary": "**Performance Summary:**\n", # Header for performance summary
    "feedback_overall_performance": "Overall: {}\n", # Overall performance level
    "feedback_domain_performance": "{}: {}\n", # Domain performance level
    "feedback_support_plan_intro": "\n**Support Plan Recommended:**\n", # Intro for support plan
    "feedback_next_steps_intro": "\n**Suggested Next Steps:**\n", # Intro for next steps
    "feedback_closing": "\nBased on these ratings, please review your updated workbook for detailed feedback and areas for development.\n\n",
    "feedback_regards": "Regards,\nSchool Leadership Team",
    "success_feedback_generated": "Feedback generated (simulated):\n\n",
    "success_feedback_log_updated": "Feedback log updated in {}",
    "error_updating_log": "Error updating feedback log in workbook:",
    "title_analytics": "Lesson Observation Analytics Dashboard", # Reverted to Lesson Observation
    "warning_no_lo_sheets_analytics": "No 'LO ' sheets found in the workbook for analytics.",
    "subheader_avg_score_overall": "Average Score per Domain (Across all observations)", # Reverted to observations
    "info_no_numeric_scores_overall": "No numeric scores found across all observations to calculate overall domain averages.", # Reverted to observations
    "subheader_data_summary": "Observation Data Summary", # Reverted to Observation
    "subheader_filter_analyze": "Filter and Analyze",
    "filter_by_school": "Filter by School",
    "filter_by_grade": "Filter by Grade",
    "filter_by_subject": "Filter by Subject",
    "option_all": "All",
    "subheader_avg_score_filtered": "Average Score per Domain (Filtered)",
    "info_no_numeric_scores_filtered": "No observations matching the selected filters contain numeric scores for domain averages.", # Reverted to observations
    "subheader_observer_distribution": "Observer Distribution (Filtered)",
    "info_no_observer_data_filtered": "No observer data found for the selected filters.",
    "info_no_observation_data_filtered": "No observation data found for the selected filters.", # Reverted to observation
    "error_loading_analytics": "Error loading or processing workbook for analytics:",
    "overall_score_label": "Overall Score:", # Label for displaying overall score
    "overall_score_value": "**{:.2f}**", # Format for displaying overall score
    "overall_score_na": "**N/A**", # Display for no numeric scores
    "arabic_toggle_label": "عرض باللغة العربية (Display in Arabic)",
    "feedback_log_sheet_name": "Feedback Log",
    "feedback_log_header": ["Sheet", "Teacher", "Email", "Observer", "School", "Subject", "Date", "Summary"],
    "download_feedback_log_csv": "📥 Download Feedback Log (CSV)",
    "error_generating_log_csv": "Error generating log CSV:",
    "download_overall_avg_csv": "📥 Download Overall Domain Averages (CSV)",
    "download_overall_avg_excel": "📥 Download Overall Domain Averages (Excel)",
    "download_filtered_avg_csv": "📥 Download Filtered Domain Averages (CSV)",
    "download_filtered_avg_excel": "📥 Download Filtered Domain Averages (Excel)",
    "download_filtered_data_csv": "📥 Download Filtered Observation Data (CSV)", # Reverted to Observation
    "download_filtered_data_excel": "📥 Download Filtered Observation Data (Excel)", # Reverted to Observation
    "label_observation_date": "Observation Date", # Reverted to Observation
    "filter_start_date": "Start Date", # New string for start date filter
    "filter_end_date": "End Date", # New string for end date filter
    "filter_teacher": "Filter by Teacher", # New string for teacher filter
    "subheader_teacher_performance": "Teacher Performance Over Time", # New subheader
    "info_select_teacher": "Select a teacher to view individual performance analytics.",
    "info_no_obs_for_teacher": "No observations found for the selected teacher within the applied filters.", # Reverted to observations
    "subheader_teacher_domain_trend": "{} Domain Performance Trend", # New subheader for teacher trend chart
    "subheader_teacher_overall_avg": "{} Overall Average Score (Filtered)", # New subheader for teacher overall avg

    # Performance Level Descriptors (English)
    "perf_level_very_weak": "Very Weak",
    "perf_level_weak": "Weak",
    "perf_level_acceptable": "Acceptable",
    "perf_level_good": "Good",
    "perf_level_excellent": "Excellent",

    # Support Plan / Next Steps Text (English) - **Customize these extensively**
    "plan_very_weak_overall": "Overall performance is Very Weak. A comprehensive support plan is required, focusing on fundamental teaching practices across multiple domains.",
    "plan_weak_overall": "Overall performance is Weak. A support plan is recommended, targeting key areas for improvement identified in the observation.",
    "plan_weak_domain": "Performance in {} is Weak. Focus on developing skills related to: {}", # Domain Name, specific elements
    "steps_acceptable_overall": "Overall performance is Acceptable. Continue to build on strengths and focus on refining practices in specific areas.",
    "steps_good_overall": "Overall performance is Good. Explore opportunities to share best practices and mentor colleagues.",
    "steps_good_domain": "Performance in {} is Good. Consider advanced strategies related to: {}", # Domain Name, specific elements
    "steps_excellent_overall": "Overall performance is Excellent. You are a role model for effective teaching. Consider leading professional development.",
    "steps_excellent_domain": "Performance in {} is Excellent. Continue to innovate and refine your practice.",
    "no_specific_plan_needed": "Performance is at an acceptable level or above. No immediate support plan required based on this observation."
}

# Placeholder Arabic strings - REPLACE THESE WITH ACTUAL TRANSLATIONS
ar_strings = {
    "page_title": "أداة التقييم للزيارات الصفية", # User's preferred translation
    "sidebar_select_page": "اختر صفحة:", # Updated translation
    "page_lesson_input": "ادخال تقييم الزيارة", # User's preferred translation
    "page_analytics": "تحليلات الزيارة", # User's preferred translation
    "title_lesson_input": "أداة إدخال زيارة صفية أسبوعية", # Updated translation
    "info_default_workbook": "استخدام مصنف القالب الافتراضي:", # Guessed translation
    "warning_default_not_found": "تحذير: لم يتم العثور على مصنف القالب الافتراضي '{}'. يرجى تحميل مصنف.", # Guessed translation
    "error_opening_default": "خطأ في فتح ملف القالب الافتراضي:", # Guessed translation
    "success_lo_sheets_found": "تم العثور على {} أوراق LO في المصنف.", # Guessed translation
    "select_sheet_or_create": "حدد ورقة LO موجودة أو أنشئ واحدة جديدة:", # Guessed translation
    "option_create_new": "إنشاء جديد", # Guessed translation
    "success_sheet_created": "تم إنشاء ورقة جديدة: {}", # Guessed translation
    "error_template_not_found": "خطأ: لم يتم العثور على ورقة القالب 'LO 1' في المصنف! لا يمكن إنشاء ورقة جديدة.", # Guessed translation
    "subheader_filling_data": "ملء البيانات لـ: {}", # Guessed translation
    "label_observer_name": "اسم المراقب", # Guessed translation
    "label_teacher_name": "اسم المعلم", # Guessed translation
    "label_teacher_email": "البريد الإلكتروني للمعلم", # Guessed translation
    "label_operator": "المشغل", # Guessed translation
    "label_school_name": "اسم المدرسة", # Guessed translation
    "label_grade": "الصف", # Guessed translation
    "label_subject": "المادة", # Guessed translation
    "label_gender": "الجنس", # Guessed translation
    "label_students": "عدد الطلاب", # Guessed translation
    "label_males": "عدد الذكور", # Guessed translation
    "label_females": "عدد الإناث", # Guessed translation
    "label_time_in": "وقت الدخول", # Guessed translation
    "label_time_out": "وقت الخروج", # Guessed translation
    "label_lesson_duration": "🕒 **مدة الدرس:** {} دقيقة — _{}_", # Guessed translation
    "duration_full_lesson": "درس كامل", # Guessed translation
    "duration_walkthrough": "جولة سريعة", # Guessed translation
    "warning_calculate_duration": "يرجى إدخال وقت الدخول ووقت الخروج لحساب المدة.", # Guessed translation
    "warning_could_not_calculate_duration": "تعذر حساب مدة الدرس:", # Guessed translation
    "label_period": "الفترة", # Guessed translation
    "label_obs_type": "نوع الزيارة", # Updated translation
    "option_individual": "فردي", # Guessed translation
    "option_joint": "مشترك", # Guessed translation
    "subheader_rubric_scores": "درجات الدليل", # Guessed translation
    "expander_rubric_descriptors": "واصفات الدليل", # Guessed translation
    "info_no_descriptors": "لا توجد واصفات دليل متاحة لهذا العنصر.", # Guessed translation
    "label_rating_for": "التقييم لـ {}", # Guessed translation
    "checkbox_send_feedback": "✉️ إرسال ملاحظات إلى المعلم", # Guessed translation
    "button_save_observation": "💾 حفظ الزيارة", # Updated translation
    "warning_fill_essential": "يرجى ملء جميع حقول المعلومات الأساسية قبل الحفظ.", # Guessed translation
    "success_data_saved": "تم حفظ بيانات الزيارة في {} في {}", # Updated translation
    "error_saving_workbook": "خطأ في حفظ المصنف:", # Guessed translation
    "download_workbook": "📥 تنزيل المصنف المحدث", # Guessed translation
    "feedback_subject": "ملاحظات الزيارة الصفية", # Updated translation
    "feedback_greeting": "عزيزي {},\n\nتم حفظ زيارتك الصفية من {}.\n\n", # Updated translation
    "feedback_observer": "المراقب: {}\n", # Guessed translation
    "feedback_duration": "المدة: {}\n", # Guessed translation
    "feedback_subject_fb": "المادة: {}\n", # Guessed translation
    "feedback_school": "المدرسة: {}\n\n", # Guessed translation
    "feedback_summary_header": "إليك ملخص لتقييماتك بناءً على الدليل:\n\n", # Guessed translation
    "feedback_domain_header": "**{}: {}**\n", # Guessed translation
    "feedback_element_rating": "- **{}:** التقييم **{}**\n", # Guessed translation
    "feedback_descriptor_for_rating": "  *واصف للتقييم {}:* {}\n", # Guessed translation
    "feedback_overall_score": "\n**متوسط الدرجة الإجمالي:** {:.2f}\n\n", # Guessed translation
    "feedback_domain_average": "  *متوسط المجال:* {:.2f}\n", # Guessed translation
    "feedback_performance_summary": "**ملخص الأداء:**\n", # Guessed translation
    "feedback_overall_performance": "الإجمالي: {}\n", # Guessed translation
    "feedback_domain_performance": "{}: {}\n", # Guessed translation
    "feedback_support_plan_intro": "\n**خطة الدعم الموصى بها:**\n", # Guessed translation
    "feedback_next_steps_intro": "\n**الخطوات التالية المقترحة:**\n", # Guessed translation
    "feedback_closing": "\nبناءً على هذه التقييمات، يرجى مراجعة المصنف المحدث للحصول على ملاحظات تفصيلية ومجالات التطوير.\n\n", # Guessed translation
    "feedback_regards": "مع التحيات,\nفريق قيادة المدرسة", # Guessed translation
    "success_feedback_generated": "تم إنشاء الملاحظات (محاكاة):\n\n", # Guessed translation
    "success_feedback_log_updated": "تم تحديث سجل الملاحظات في {}", # Guessed translation
    "error_updating_log": "خطأ في تحديث سجل الملاحظات في المصنف:", # Guessed translation
    "title_analytics": "لوحة تحليلات الزيارة الصفية", # Updated translation
    "warning_no_lo_sheets_analytics": "لم يتم العثور على أوراق 'LO ' في المصنف للتحليلات.", # Guessed translation
    "subheader_avg_score_overall": "متوسط الدرجة لكل مجال (عبر جميع الزيارات)", # Updated translation
    "info_no_numeric_scores_overall": "لم يتم العثور على درجات رقمية عبر جميع الزيارات لحساب متوسطات المجال الإجمالية.", # Updated translation
    "subheader_data_summary": "ملخص بيانات الزيارة", # Updated translation
    "subheader_filter_analyze": "تصفية وتحليل", # Guessed translation
    "filter_by_school": "تصفية حسب المدرسة", # Guessed translation
    "filter_by_grade": "تصفية حسب الصف", # Guessed translation
    "filter_by_subject": "تصفية حسب المادة", # Guessed translation
    "option_all": "الكل", # Guessed translation
    "subheader_avg_score_filtered": "متوسط الدرجة لكل مجال (مصفى)", # Guessed translation
    "info_no_numeric_scores_filtered": "لا توجد زيارات مطابقة للمرشحات المحددة تحتوي على درجات رقمية لمتوسطات المجال.", # Updated translation
    "subheader_observer_distribution": "توزيع المراقبين (مصفى)", # Guessed translation
    "info_no_observer_data_filtered": "لم يتم العثور على بيانات المراقب للمرشحات المحددة.", # Guessed translation
    "info_no_observation_data_filtered": "لم يتم العثور على بيانات الزيارة للمرشحات المحددة.", # Updated translation
    "error_loading_analytics": "خطأ في تحميل أو معالجة المصنف للتحليلات:", # Guessed translation
    "overall_score_label": "النتيجة الإجمالية:", # Guessed translation
    "overall_score_value": "**{:.2f}**", # Guessed translation
    "overall_score_na": "**غير متوفر**", # Guessed translation
    "arabic_toggle_label": "عرض باللغة العربية (Display in Arabic)", # Keep English part as requested
    "feedback_log_sheet_name": "سجل الملاحظات", # Guessed translation
    "feedback_log_header": ["الورقة", "المعلم", "البريد الإلكتروني", "المراقب", "المدرسة", "المادة", "التاريخ", "الملخص"], # Guessed translation
    "download_feedback_log_csv": "📥 تنزيل سجل الملاحظات (CSV)", # Guessed translation
    "error_generating_log_csv": "خطأ في إنشاء سجل الملاحظات CSV:", # Guessed translation
    "download_overall_avg_csv": "📥 تنزيل متوسطات المجال الإجمالية (CSV)", # Guessed translation
    "download_overall_avg_excel": "📥 تنزيل متوسطات المجال الإجمالية (Excel)", # Guessed translation
    "download_filtered_avg_csv": "📥 تنزيل متوسطات المجال المصفاة (CSV)", # Guessed translation
    "download_filtered_avg_excel": "📥 تنزيل متوسطات المجال المصفاة (Excel)", # Guessed translation
    "download_filtered_data_csv": "📥 تنزيل بيانات الزيارة المصفاة (CSV)", # Updated translation
    "download_filtered_data_excel": "📥 تنزيل بيانات الزيارة المصفاة (Excel)", # Updated translation
    "label_observation_date": "تاريخ الزيارة", # Updated translation
    "filter_start_date": "تاريخ البدء", # Guessed translation
    "filter_end_date": "تاريخ الانتهاء", # Guessed translation
    "filter_teacher": "تصفية حسب المعلم", # Guessed translation
    "subheader_teacher_performance": "أداء المعلم بمرور الوقت", # Guessed translation
    "info_select_teacher": "حدد معلمًا لعرض تحليلات الأداء الفردي.", # Guessed translation
    "info_no_obs_for_teacher": "لم يتم العثور على زيارات للمعلم المحدد ضمن المرشحات المطبقة.", # Updated translation
    "subheader_teacher_domain_trend": "اتجاه أداء مجال {}", # Guessed translation
    "subheader_teacher_overall_avg": "متوسط الدرجة الإجمالي لـ {} (مصفى)", # Guessed translation

    # Performance Level Descriptors (Arabic) - **Translate these**
    "perf_level_very_weak": "ضعيف جداً", # Guessed translation
    "perf_level_weak": "ضعيف", # Guessed translation
    "perf_level_acceptable": "مقبول", # Guessed translation
    "perf_level_good": "جيد", # Guessed translation
    "perf_level_excellent": "ممتاز", # Guessed translation

    # Support Plan / Next Steps Text (Arabic) - **Translate and Customize these extensively**
    "plan_very_weak_overall": "الأداء الإجمالي ضعيف جداً. تتطلب خطة دعم شاملة، تركز على الممارسات التعليمية الأساسية عبر مجالات متعددة.", # Guessed translation
    "plan_weak_overall": "الأداء الإجمالي ضعيف. يوصى بخطة دعم، تستهدف المجالات الرئيسية للتحسين المحددة في الملاحظة.", # Guessed translation
    "plan_weak_domain": "الأداء في {} ضعيف. ركز على تطوير المهارات المتعلقة بـ: {}", # Guessed translation
    "steps_acceptable_overall": "الأداء الإجمالي مقبول. استمر في البناء على نقاط القوة وركز على تحسين الممارسات في مجالات محددة.", # Guessed translation
    "steps_good_overall": "الأداء الإجمالي جيد. استكشف فرص مشاركة أفضل الممارسات وتوجيه الزملاء.", # Guessed translation
    "steps_good_domain": "الأداء في {} جيد. فكر في استراتيجيات متقدمة تتعلق بـ: {}", # Guessed translation
    "steps_excellent_overall": "الأداء الإجمالي ممتاز. أنت نموذج يحتذى به في التدريس الفعال. فكر في قيادة التطوير المهني.", # Guessed translation
    "steps_excellent_domain": "الأداء في {} ممتاز. استمر في الابتكار وتحسين ممارستك.", # Guessed translation
    "no_specific_plan_needed": "الأداء عند مستوى مقبول أو أعلى. لا توجد خطة دعم فورية مطلوبة بناءً على هذه الملاحظة.", # Guessed translation
}

# --- Function to get strings based on language toggle ---
def get_strings(arabic_mode):
    return ar_strings if arabic_mode else en_strings

# --- Function to determine performance level based on score ---
def get_performance_level(score, strings):
    if score is None or (isinstance(score, float) and math.isnan(score)):
        return strings["overall_score_na"] # Or a specific string for no score
    # Ensure score is treated as a number for comparison
    try:
        numeric_score = float(score)
        if numeric_score >= 5.5: # Example thresholds - Adjust as needed
            return strings["perf_level_excellent"]
        elif numeric_score >= 4.5:
            return strings["perf_level_good"]
        elif numeric_score >= 3.5:
            return strings["perf_level_acceptable"]
        elif numeric_score >= 2.5:
            return strings["perf_level_weak"]
        else:
            return strings["perf_level_very_weak"]
    except (ValueError, TypeError):
        return strings["overall_score_na"] # Handle cases where score is not a valid number


# --- Function to generate PDF ---
def generate_observation_pdf(data, feedback_content, strings, rubric_domains_structure):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()

    # Custom styles
    styles.add(ParagraphStyle(name='Heading1Centered', alignment=1, fontSize=16, spaceAfter=14, bold=1))
    styles.add(ParagraphStyle(name='Heading2', fontSize=12, spaceAfter=10, bold=1))
    styles.add(ParagraphStyle(name='Normal', fontSize=10, spaceAfter=6))
    styles.add(ParagraphStyle(name='RubricDescriptor', fontSize=9, spaceAfter=4, leftIndent=18)) # Indent descriptors
    styles.add(ParagraphStyle(name='RubricDomainHeading', fontSize=11, spaceAfter=8, bold=1)) # Style for domain headings in PDF
    styles.add(ParagraphStyle(name='RubricElementRating', fontSize=10, spaceAfter=4, leftIndent=10)) # Style for element rating in PDF


    story = []

    # --- Add School Logo ---
    school_name = data.get("School", "Default")
    logo_path = LOGO_PATHS.get(school_name, LOGO_PATHS["Default"])

    if os.path.exists(logo_path):
        try:
            # Adjust width and height as needed
            img = Image(logo_path, width=1.5*inch, height=0.75*inch)
            img.hAlign = 'CENTER' # Center the logo
            story.append(img)
            story.append(Spacer(1, 0.2*inch)) # Add space after the logo
        except Exception as e:
            st.warning(f"Could not add logo for {school_name}: {e}")
            # Optionally add a text placeholder if logo fails
            # story.append(Paragraph(f"[{school_name} Logo Placeholder]", styles['Normal']))
    else:
        st.warning(f"Logo file not found for {school_name} at {logo_path}. Using text title.")
        # Fallback to just the title if logo file is missing
        story.append(Paragraph(strings["page_title"], styles['Heading1Centered']))
        story.append(Spacer(1, 0.2*inch))


    # Basic Information Table
    basic_info_data = [
        [strings["label_observer_name"] + ":", data.get("Observer Name", "")],
        [strings["label_teacher_name"] + ":", data.get("Teacher", "")],
        [strings["label_teacher_email"] + ":", data.get("Teacher Email", "")],
        [strings["label_operator"] + ":", data.get("Operator", "")],
        [strings["label_school_name"] + ":", data.get("School", "")],
        [strings["label_grade"] + ":", data.get("Grade", "")],
        [strings["label_subject"] + ":", data.get("Subject", "")],
        [strings["label_gender"] + ":", data.get("Gender", "")],
        [strings["label_students"] + ":", data.get("Students", "")],
        [strings["label_males"] + ":", data.get("Males", "")],
        [strings["label_females"] + ":", data.get("Females", "")],
        [strings["label_observation_date"] + ":", data.get("Observation Date", "")],
        [strings["label_time_in"] + ":", data.get("Time In", "")],
        [strings["label_time_out"] + ":", data.get("Time Out", "")],
        [strings["label_lesson_duration"] + ":", data.get("Duration", "")], # Using the duration label
        [strings["label_period"] + ":", data.get("Period", "")],
        [strings["label_obs_type"] + ":", data.get("Observation Type", "")],
        [strings["overall_score_label"] + ":", data.get("Overall Score", strings["overall_score_na"])] # Include Overall Score
    ]

    table_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
    ])

    # Need to handle potential None values in data
    cleaned_basic_info_data = [[item[0], str(item[1]) if item[1] is not None else "N/A"] for item in basic_info_data]


    table = Table(cleaned_basic_info_data, colWidths=[2*inch, 4*inch])
    table.setStyle(table_style)
    story.append(table)
    story.append(Spacer(1, 0.2*inch))

    # Rubric Scores - Including detailed scores and descriptors
    story.append(Paragraph(strings["subheader_rubric_scores"], styles['Heading2']))

    element_ratings_data = data.get("element_ratings", {})
    domain_avg_scores_data = data.get("domain_avg_scores", {})

    for domain, (start_cell, count) in rubric_domains_structure.items():
        # Get domain title (assuming English title is sufficient for structure in PDF)
        # If you need localized domain titles in PDF, you'd need to pass them or look them up
        # For accurate domain titles in PDF, you should pass the actual titles from your Excel/app
        # For now, using domain name as a placeholder in PDF
        domain_title_display = domain # Use domain name for now

        story.append(Paragraph(f"{domain}: {domain_title_display}", styles['RubricDomainHeading']))

        # Add domain average to PDF
        avg_score = domain_avg_scores_data.get(domain)
        if avg_score is not None:
             story.append(Paragraph(f"  Domain Average: {avg_score:.2f}", styles['Normal']))
        else:
             story.append(Paragraph(f"  Domain Average: {strings['overall_score_na']}", styles['Normal']))


        for i in range(count):
            element_key = f"{domain}_{i}"
            element_info = element_ratings_data.get(element_key)

            if element_info:
                label_en = element_info.get('label_en', f"Element {domain[-1]}.{i+1}")
                rating = element_info.get('rating', 'N/A')
                description_en = element_info.get('description_en', '') # Get the full formatted descriptor text

                story.append(Paragraph(f"- <b>{label_en}:</b> Rating <b>{rating}</b>", styles['RubricElementRating']))

                # Include the full descriptor text if available
                if description_en:
                     # Convert the markdown-like descriptor text to ReportLab flowables
                     descriptor_paragraphs = description_en.split('\n\n')
                     for desc_para in descriptor_paragraphs:
                          if desc_para.strip():
                               desc_para = desc_para.replace('**', '<b>').replace('**', '</b>')
                               story.append(Paragraph(desc_para.replace('\n', '<br/>'), styles['RubricDescriptor']))
                         # story.append(Spacer(1, 0.05*inch)) # Smaller space between descriptor paragraphs
                story.append(Spacer(1, 0.1*inch)) # Space after each element


        story.append(Spacer(1, 0.2*inch)) # Space after each domain


    # Feedback Content
    story.append(Paragraph("Feedback Report:", styles['Heading2']))
    # The feedback_content string already contains formatting (like **, \n)
    # We need to convert this markdown-like text to ReportLab flowables
    # This is a simplified conversion; a full markdown parser would be more robust
    feedback_paragraphs = feedback_content.split('\n\n') # Split by double newline for paragraphs
    for para in feedback_paragraphs:
        if para.strip(): # Avoid empty paragraphs
            # Simple bold conversion
            para = para.replace('**', '<b>').replace('**', '</b>')
            story.append(Paragraph(para.replace('\n', '<br/>'), styles['Normal'])) # Replace single newlines with breaks
        story.append(Spacer(1, 0.1*inch)) # Add space between paragraphs


    # Build the PDF
    try:
        doc.build(story)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"Error generating PDF: {e}")
        return None


# --- Streamlit App Layout ---
# Add Arabic toggle early to affect language throughout the app
arabic_mode = st.sidebar.toggle(en_strings["arabic_toggle_label"], False)
strings = get_strings(arabic_mode)

# Set page config using the selected language
# MOVED TO TOP

# Sidebar page selection
page = st.sidebar.selectbox(strings["sidebar_select_page"], [strings["page_lesson_input"], strings["page_analytics"]])

uploaded_file = None
DEFAULT_FILE = "Teaching Rubric Tool_WeekTemplate.xlsx"
# Check if the default file exists and is readable before trying to open
if os.path.exists(DEFAULT_FILE):
    try:
        uploaded_file = open(DEFAULT_FILE, "rb")
        st.info(strings["info_default_workbook"].format(DEFAULT_FILE))
    except Exception as e:
        st.error(strings["error_opening_default"].format(e))
        uploaded_file = None # Ensure uploaded_file is None if opening fails
else:
    st.warning(strings["warning_default_not_found"].format(DEFAULT_FILE))


if page == strings["page_lesson_input"]:
    st.title(strings["title_lesson_input"])

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
            st.success(strings["success_lo_sheets_found"].format(len(lo_sheets)))

            selected_option = st.selectbox(strings["select_sheet_or_create"], [strings["option_create_new"]] + lo_sheets)

            if selected_option == strings["option_create_new"]:
                next_index = 1
                while f"LO {next_index}" in wb.sheetnames:
                    next_index += 1
                sheet_name = f"LO {next_index}"
                # Ensure the template sheet "LO 1" exists before copying
                if "LO 1" in wb.sheetnames:
                     wb.copy_worksheet(wb["LO 1"]).title = sheet_name
                     st.success(strings["success_sheet_created"].format(sheet_name))
                else:
                     st.error(strings["error_template_not_found"])
                     st.stop() # Stop execution if template is missing

            else:
                sheet_name = selected_option

            ws = wb[sheet_name]
            st.subheader(strings["subheader_filling_data"].format(sheet_name))

            observer = st.text_input(strings["label_observer_name"])
            teacher = st.text_input(strings["label_teacher_name"])
            teacher_email = st.text_input(strings["label_teacher_email"])
            operator = st.selectbox(strings["label_operator"], sorted(["Taaleem", "Al Dar", "New Century Education", "Bloom"]))

            # School options - You might need to translate these names or provide Arabic options
            school_options = {
                "New Century Education": ["Al Bayan School", "Al Bayraq School", "Al Dhaher School", "Al Hosoon School", "Al Mutanabi School", "Al Nahdha School", "Jern Yafoor School", "Maryam Bint Omran School"],
                "Taaleem": ["Al Ahad Charter School", "Al Azm Charter School", "Al Riyadh Charter School", "Al Majd Charter School", "Al Qeyam Charter School", "Al Nayfa Charter Kindergarten", "Al Salam Charter School", "Al Walaa Charter Kindergarten", "Al Forsan Charter Kindergarten", "Al Wafaa Charter Kindergarten", "Al Watan Charter School"],
                "Al Dar": ["Al Ghad Charter School", "Al Mushrif Charter Kindergarten", "Al Danah Charter School", "Al Rayaheen Charter School", "Al Rayana Charter School", "Al Qurm Charter School", "Mubarak Bin Mohammed Charter School (Cycle 2 & 3)"],
                "Bloom": ["Al Ain Charter School", "Al Dana Charter School", "Al Ghadeer Charter School", "Al Hili Charter School", "Al Manhal Charter School", "Al Qattara Charter School", "Al Towayya Charter School", "Jabel Hafeet Charter School"]
            }

            school_list = sorted(school_options.get(operator, []))
            school = st.selectbox(strings["label_school_name"], school_list)
            # Grade options - You might need to translate these
            grade = st.selectbox(strings["label_grade"], [f"Grade {i}" for i in range(1, 13)] + ["K1", "K2"])
            # Subject options - You might need to translate these
            subject = st.selectbox(strings["label_subject"], ["Math", "English", "Arabic", "Science", "Islamic", "Social Studies"])
            # Gender options - You might need to translate these
            gender = st.selectbox(strings["label_gender"], ["Male", "Female", "Mixed"])
            students = st.text_input(strings["label_students"])
            males = st.text_input(strings["label_males"])
            females = st.text_input(strings["label_females"])
            time_in = st.time_input(strings["label_time_in"])
            time_out = st.time_input(strings["label_time_out"])
            # New date input field
            observation_date = st.date_input(strings["label_observation_date"], datetime.now().date())


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
                    duration_label = strings["duration_full_lesson"] if minutes >= 40 else strings["duration_walkthrough"]
                    st.markdown(strings["label_lesson_duration"].format(minutes, duration_label))
                else:
                     st.warning(strings["warning_calculate_duration"])
            except Exception as e:
                st.warning(strings["warning_could_not_calculate_duration"].format(e))


            period = st.selectbox(strings["label_period"], [f"Period {i}" for i in range(1, 9)])
            obs_type = st.selectbox(strings["label_obs_type"], [strings["option_individual"], strings["option_joint"]])

            rubric_domains = {
                "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
            }

            st.markdown("---")
            st.subheader(strings["subheader_rubric_scores"])

            domain_colors = ["#e6f2ff", "#fff2e6", "#e6ffe6", "#f9e6ff", "#ffe6e6", "#f0f0f5", "#e6f9ff", "#f2ffe6", "#ffe6f2"]
            all_element_ratings = {} # To store ratings for feedback generation and overall score
            domain_element_ratings_list = {domain: [] for domain in rubric_domains.keys()} # To store numeric ratings per domain

            for idx, (domain, (start_cell, count)) in enumerate(rubric_domains.items()):
                background = domain_colors[idx % len(domain_colors)]
                row = int(start_cell[1:])
                # Get domain title from Excel (assuming English in Excel)
                domain_title_en = ws[f'A{row}'].value or domain
                # If you have Arabic titles in Excel, you'd need logic here to pick based on language
                # For now, using English title from Excel
                domain_title_display = domain_title_en # Use English title from Excel


                st.markdown(f"""
                <div style='background-color:{background};padding:12px;border-radius:10px;margin-bottom:5px;'>
                <h4 style='margin-bottom:5px;'>{domain}: {domain_title_display}</h4>
                </div>
                """, unsafe_allow_html=True)

                col = start_cell[0]

                for i in range(count):
                    element_row = row + i
                    # Get the element label from column B (assuming English in Excel)
                    label_en = ws[f"B{element_row}"].value or f"Element {domain[-1]}.{i+1}"
                     # If you have Arabic labels in Excel, you'd need logic here
                    label_display = label_en # Use English label from Excel

                    # Collect rubric descriptors from columns C to H (assuming English in Excel)
                    rubric_descriptors_en = [ws[f"{chr(ord('C')+j)}{element_row}"].value for j in range(6)]
                    # Filter out None or empty strings and format for display
                    formatted_rubric_text_display = "\n\n".join([f"**{6-j}:** {desc}" for j, desc in enumerate(rubric_descriptors_en) if desc])

                    st.markdown(f"**{label_display}**")
                    with st.expander(strings["expander_rubric_descriptors"]):
                        if formatted_rubric_text_display:
                            st.markdown(formatted_rubric_text_display)
                        else:
                            st.info(strings["info_no_descriptors"])

                    # Get the rating from the user
                    rating = st.selectbox(strings["label_rating_for"].format(label_display), [6, 5, 4, 3, 2, 1, "NA"], key=f"{sheet_name}_{domain}_{i}")

                    # Store the rating and description for feedback generation and overall score
                    all_element_ratings[f"{domain}_{i}"] = {"label_en": label_en, "rating": rating, "description_en": formatted_rubric_text_display} # Store English label and description

                    # If the rating is numeric, add it to the domain's list
                    if isinstance(rating, int):
                         domain_element_ratings_list[domain].append(rating)

                    # Write the rating to the Excel sheet
                    ws[f"{col}{element_row}"] = rating

            # --- Calculate and Display Overall and Domain Scores ---
            numeric_ratings = [
                item['rating'] for item in all_element_ratings.values()
                if isinstance(item['rating'], int) # Only include numeric ratings
            ]

            overall_score = None
            if numeric_ratings:
                overall_score = statistics.mean(numeric_ratings)

            domain_avg_scores = {}
            for domain, ratings in domain_element_ratings_list.items():
                 if ratings:
                     domain_avg_scores[domain] = statistics.mean(ratings)
                 else:
                     domain_avg_scores[domain] = None # Or 0, depending on how you want to represent domains with no numeric scores

            st.markdown("---") # Separator before scores
            st.subheader(strings["overall_score_label"])
            if overall_score is not None:
                st.markdown(strings["overall_score_value"].format(overall_score))
            else:
                st.markdown(strings["overall_score_na"])

            st.subheader("Domain Average Scores:") # New subheader for domain averages
            for domain, avg_score in domain_avg_scores.items():
                 domain_title_en = ws[f'A{int(rubric_domains[domain][0][1:])}'].value or domain
                 if avg_score is not None:
                      st.markdown(f"- **{domain}:** {domain_title_en}: {avg_score:.2f}")
                 else:
                      st.markdown(f"- **{domain}:** {domain_title_en}: {strings['overall_score_na']}")


            st.markdown("---") # Separator after scores


            send_feedback = st.checkbox(strings["checkbox_send_feedback"])

            if st.button(strings["button_save_observation"]):
                # Ensure essential fields are filled before saving (optional but good practice)
                if not all([observer, teacher, school, grade, subject, students, males, females, time_in, time_out, observation_date]):
                     st.warning(strings["warning_fill_essential"])
                else: # Proceed only if essential fields are filled
                    ws["Z1"] = strings["label_observer_name"]; ws["AA1"] = observer
                    ws["Z2"] = strings["label_teacher_name"]; ws["AA2"] = teacher
                    ws["Z3"] = strings["label_obs_type"]; ws["AA3"] = obs_type
                    ws["Z4"] = strings["label_operator"]; ws["AA4"] = operator
                    ws["Z5"] = strings["label_school_name"]; ws["AA5"] = school
                    ws["Z6"] = strings["label_subject"]; ws["AA6"] = subject
                    ws["Z7"] = strings["label_grade"]; ws["AA7"] = grade
                    ws["Z8"] = strings["label_gender"]; ws["AA8"] = gender
                    ws["Z9"] = strings["label_students"]; ws["AA9"] = students # Consider converting to int if needed elsewhere
                    ws["Z10"] = strings["label_males"]; ws["AA10"] = males # Consider converting to int if needed elsewhere
                    ws["Z11"] = strings["label_females"]; ws["AA11"] = females # Consider converting to int if needed elsewhere
                    ws["Z12"] = strings["label_lesson_duration"]; ws["AA12"] = duration_label
                    ws["Z13"] = strings["label_time_in"]
                    # Check if time_in is not None before formatting
                    ws["AA13"] = time_in.strftime("%H:%M") if time_in else "N/A"
                    ws["Z14"] = strings["label_time_out"]
                     # Check if time_out is not None before formatting
                    ws["AA14"] = time_out.strftime("%H:%M") if time_out else "N/A"
                    # Save the overall score to the sheet
                    ws["Z15"] = strings["overall_score_label"]
                    ws["AA15"] = overall_score if overall_score is not None else "N/A"
                    # Save the observation date to the sheet
                    ws["Z16"] = strings["label_observation_date"] # New row for date label
                    ws["AA16"] = observation_date # New row for date value (saved as datetime.date object)


                    # Save domain average scores to the sheet (Optional - choose columns)
                    # Example: Save to columns AB onwards, starting from row 15
                    domain_avg_start_col_idx = ord('AB') # ASCII value of 'AB'
                    domain_avg_start_row = 15
                    for domain_idx, (domain, (start_cell, count)) in enumerate(rubric_domains.items()):
                         col_letter = chr(domain_avg_start_col_idx + domain_idx) # AB, AC, AD...
                         ws[f"{col_letter}{domain_avg_start_row}"] = f"{domain} Avg"
                         ws[f"{col_letter}{domain_avg_start_row + 1}"] = domain_avg_scores.get(domain, "N/A") # Save avg score below label


                    save_path = f"updated_{sheet_name}.xlsx"
                    try:
                        wb.save(save_path)
                        st.success(strings["success_data_saved"].format(sheet_name, save_path))
                        # The Excel download button is here
                        with open(save_path, "rb") as f:
                            st.download_button(strings["download_workbook"], f, file_name=save_path)
                        # os.remove(save_path) # Consider keeping the file for a bit or saving to a specific temp directory

                    except Exception as e:
                         st.error(f"Error saving workbook: {e}") # More specific error message


                    # Generate and send feedback
                    if send_feedback and teacher_email:
                        feedback_content = strings["feedback_greeting"].format(teacher, observation_date.strftime('%Y-%m-%d')) # Use observation_date
                        feedback_content += strings["feedback_observer"].format(observer)
                        feedback_content += strings["feedback_duration"].format(duration_label)
                        feedback_content += strings["feedback_subject_fb"].format(subject)
                        feedback_content += strings["feedback_school"].format(school)
                        feedback_content += strings["feedback_summary_header"]

                        # Add rubric scores and descriptions to feedback
                        for domain, (start_cell, count) in rubric_domains.items():
                             # Get domain title from Excel (assuming English)
                             domain_title_en = ws[f'A{int(start_cell[1:])}'].value or domain
                             feedback_content += strings["feedback_domain_header"].format(domain, domain_title_en) # Use English title in feedback

                             # Add domain average to feedback
                             if domain in domain_avg_scores and domain_avg_scores[domain] is not None:
                                 feedback_content += strings["feedback_domain_average"].format(domain_avg_scores[domain])
                             else:
                                 feedback_content += strings["feedback_domain_average"].format(strings['overall_score_na'])


                             for i in range(count):
                                 element_key = f"{domain}_{i}"
                                 element_info = all_element_ratings.get(element_key)
                                 if element_info:
                                     rating = element_info['rating']
                                     label_en = element_info['label_en'] # Use English label in feedback

                                     feedback_content += strings["feedback_element_rating"].format(label_en, rating)

                                     # Optional: Include the descriptor for the given rating level
                                     if rating != "NA" and isinstance(rating, int):
                                         try:
                                             # Find the descriptor for the specific rating
                                             # Assuming rubric_descriptors were collected in order 6, 5, 4, 3, 2, 1
                                             descriptor_index = 6 - rating
                                             # Re-read descriptors to get the specific one from Excel (assuming English)
                                             element_row = int(start_cell[1:]) + i
                                             specific_descriptor_en = ws[f"{chr(ord('C')+descriptor_index)}{element_row}"].value
                                             if specific_descriptor_en:
                                                 feedback_content += strings["feedback_descriptor_for_rating"].format(rating, specific_descriptor_en)
                                         except (IndexError, ValueError):
                                             pass # Handle potential issues with index or rating value

                             feedback_content += "\n" # Add space between domains

                        # Add overall score to feedback
                        if overall_score is not None:
                            feedback_content += strings["feedback_overall_score"].format(overall_score)
                        else:
                             feedback_content += strings["overall_score_label"] + strings["overall_score_na"] + "\n\n"

                        # --- Add Performance Summary and Support/Next Steps ---
                        feedback_content += strings["feedback_performance_summary"]
                        feedback_content += strings["feedback_overall_performance"].format(get_performance_level(overall_score, strings))

                        # Add performance level for each domain
                        for domain, avg_score in domain_avg_scores.items():
                            domain_title_en = ws[f'A{int(rubric_domains[domain][0][1:])}'].value or domain
                            feedback_content += strings["feedback_domain_performance"].format(domain_title_en, get_performance_level(avg_score, strings))

                        # Add Support Plan or Next Steps based on Overall Score
                        if overall_score is not None:
                             if overall_score < 3.5: # Example threshold for Weak/Very Weak
                                 feedback_content += strings["feedback_support_plan_intro"]
                                 if overall_score < 2.5: # Example threshold for Very Weak
                                      feedback_content += strings["plan_very_weak_overall"] + "\n"
                                 else: # Weak
                                      feedback_content += strings["plan_weak_overall"] + "\n"

                                 # Suggest domains to focus on if weak
                                 weak_domains = [d for d, avg in domain_avg_scores.items() if avg is not None and avg < 3.5] # Example threshold
                                 if weak_domains:
                                      feedback_content += "Areas for focus include: " + ", ".join([ws[f'A{int(rubric_domains[d][0][1:])}'].value or d for d in weak_domains]) + "\n"

                                 # Placeholder for more specific AI-generated support plan items
                                 feedback_content += "\n[Placeholder: Specific support plan items to be discussed with school leadership or generated with AI assistance.]\n"

                             elif overall_score >= 3.5: # Example threshold for Acceptable and above
                                 feedback_content += strings["feedback_next_steps_intro"]
                                 if overall_score >= 5.5: # Excellent
                                     feedback_content += strings["steps_excellent_overall"] + "\n"
                                 elif overall_score >= 4.5: # Good
                                     feedback_content += strings["steps_good_overall"] + "\n"
                                 else: # Acceptable
                                     feedback_content += strings["steps_acceptable_overall"] + "\n"

                                 # Suggest domains of strength if good/excellent
                                 strong_domains = [d for d, avg in domain_avg_scores.items() if avg is not None and avg >= 4.5] # Example threshold
                                 if strong_domains:
                                      feedback_content += "Areas of strength include: " + ", ".join([ws[f'A{int(rubric_domains[d][0][1:])}'].value or d for d in strong_domains]) + "\n"

                                 # Placeholder for more specific AI-generated next steps
                                 feedback_content += "\n[Placeholder: Specific next steps to be discussed with school leadership or generated with AI assistance.]\n"

                             else:
                                  feedback_content += strings["no_specific_plan_needed"] + "\n"


                        feedback_content += strings["feedback_closing"]
                        feedback_content += strings["feedback_regards"]


                        st.success(strings["success_feedback_generated"] + feedback_content)

                        # --- Generate and Download PDF ---
                        # Prepare data for PDF generation
                        pdf_data = {
                            "Observer Name": observer,
                            "Teacher": teacher,
                            "Teacher Email": teacher_email,
                            "Operator": operator,
                            "School": school,
                            "Grade": grade,
                            "Subject": subject,
                            "Gender": gender,
                            "Students": students,
                            "Males": males,
                            "Females": females,
                            "Observation Date": observation_date.strftime('%Y-%m-%d') if observation_date else "N/A",
                            "Time In": time_in.strftime("%H:%M") if time_in else "N/A",
                            "Time Out": time_out.strftime("%H:%M") if time_out else "N/A",
                            "Duration": duration_label,
                            "Period": period,
                            "Observation Type": obs_type,
                            "Overall Score": overall_score, # Pass the calculated score
                            "element_ratings": all_element_ratings, # Pass element ratings for PDF
                            "domain_avg_scores": domain_avg_scores # Pass domain avg scores for PDF
                        }
                        # You also need to pass the rubric_domains structure
                        try:
                            pdf_buffer = generate_observation_pdf(pdf_data, feedback_content, strings, rubric_domains)

                            if pdf_buffer:
                                st.download_button(
                                    label="📥 Download Observation PDF",
                                    data=pdf_buffer,
                                    file_name=f"{sheet_name}_Observation_Report.pdf",
                                    mime="application/pdf"
                                )
                        except Exception as e:
                            st.error(f"Error generating PDF: {e}") # More specific error message


                        # Feedback log to sheet
                        try:
                            if strings["feedback_log_sheet_name"] not in wb.sheetnames:
                                log_ws = wb.create_sheet(strings["feedback_log_sheet_name"])
                                log_ws.append(strings["feedback_log_header"])
                            else:
                                log_ws = wb[strings["feedback_log_sheet_name"]]

                            # Append log entry
                            log_ws.append([
                                sheet_name, teacher, teacher_email, observer, school, subject,
                                observation_date.strftime("%Y-%m-%d"), # Use observation_date for log
                                feedback_content[:500] + ("..." if len(feedback_content) > 500 else "") # Truncate summary
                            ])

                            # Save the workbook again to include the log entry
                            # This save is crucial for the log to persist
                            wb.save(save_path)
                            st.success(strings["success_feedback_log_updated"].format(save_path))

                        except Exception as e:
                            st.error(f"Error updating feedback log in workbook: {e}") # More specific error message

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
                        #         st.download_button(strings["download_feedback_log_csv"], f, file_name=log_csv_path)
                        #     # os.remove(log_csv_path) # Clean up the temporary file
                        # except Exception as e:
                        #      st.error(strings["error_generating_log_csv"].format(e))

        except Exception as e:
            st.error(f"Error loading or processing workbook for input: {e}") # More specific error message in input section


elif page == strings["page_analytics"]:
    st.title(strings["title_analytics"])

    if uploaded_file:
        try:
            # Use data_only=True to get calculated values from the Excel file
            wb = load_workbook(uploaded_file, data_only=True)
            sheets = [s for s in wb.sheetnames if s.startswith("LO ")]

            if not sheets:
                st.warning(strings["warning_no_lo_sheets_analytics"])
            else:
                # Define rubric domains again for analytics page calculations
                rubric_domains = {
                    "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                    "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
                }
                all_domains_list = [f"Domain {i}" for i in range(1, 10)]


                metadata = []
                all_observations_domain_scores = [] # To store domain scores for each observation

                for sheet in sheets:
                    ws = wb[sheet]
                    # Read the observation date from the new cell (AA16)
                    observation_date_value = ws["AA16"].value
                    # Attempt to convert to date object, handle errors
                    obs_date = None
                    if isinstance(observation_date_value, datetime):
                        obs_date = observation_date_value.date()
                    elif isinstance(observation_date_value, date):
                         obs_date = observation_date_value
                    elif isinstance(observation_date_value, str):
                         try:
                             obs_date = datetime.strptime(observation_date_value, "%Y-%m-%d").date() # Assuming %Y-%m-%d format if saved as string
                         except (ValueError, TypeError):
                              pass # Ignore if string format is unexpected


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
                        "Overall Score": ws["AA15"].value, # Added Overall Score from saved data
                        "Observation Date": obs_date # Add the observation date
                     }
                    metadata.append(row_info)

                    # Calculate average score for each domain in this sheet and store
                    sheet_domain_avg = {} # Store average for each domain in the current sheet
                    for domain, (start_cell, count) in rubric_domains.items():
                        col = start_cell[0]
                        row = int(start_cell[1:])
                        domain_element_ratings = [] # Ratings for elements within THIS domain in THIS sheet
                        for j in range(count):
                            cell_value = ws[f"{col}{row + j}"].value
                            try:
                                # Attempt to convert to float, handles ints and floats. Exclude NaN.
                                rating = float(cell_value)
                                if not math.isnan(rating): # Check if it's not NaN
                                    domain_element_ratings.append(rating)
                            except (ValueError, TypeError):
                                # Ignore non-numeric values like "NA", None, or errors
                                pass

                        if domain_element_ratings:
                            # Calculate average for this domain in THIS sheet
                            sheet_domain_avg[domain] = statistics.mean(domain_element_ratings)
                        else:
                             sheet_domain_avg[domain] = None # Store None if no numeric ratings for the domain


                    # Store domain averages for this observation, linked to the sheet name and date
                    observation_domain_data = {"Sheet": sheet, "Observation Date": obs_date}
                    observation_domain_data.update(sheet_domain_avg)
                    all_observations_domain_scores.append(observation_domain_data)


                # Create DataFrames
                df_meta = pd.DataFrame(metadata)
                df_domain_scores_all = pd.DataFrame(all_observations_domain_scores)

                # Ensure 'Observation Date' is datetime type for sorting and filtering
                df_meta['Observation Date'] = pd.to_datetime(df_meta['Observation Date'], errors='coerce').dt.date
                df_domain_scores_all['Observation Date'] = pd.to_datetime(df_domain_scores_all['Observation Date'], errors='coerce').dt.date

                # Sort by date for trend analysis
                df_meta = df_meta.sort_values(by="Observation Date")
                df_domain_scores_all = df_domain_scores_all.sort_values(by="Observation Date")


                # Calculate overall averages from df_domain_scores_all for the overall chart
                # Filter out rows with no valid date before calculating overall averages
                df_domain_scores_valid_dates = df_domain_scores_all.dropna(subset=['Observation Date'])
                avg_scores = {
                    domain: round(df_domain_scores_valid_dates[domain].mean(), 2) if not df_domain_scores_valid_dates[domain].isnull().all() else 0
                    for domain in all_domains_list
                }


                # Ensure all domains are included, using 0 if no scores were collected
                overall_avg_data = [{"Domain": d, "Average Score": avg_scores.get(d, 0)} for d in all_domains_list]
                df_avg = pd.DataFrame(overall_avg_data)

                st.subheader(strings["subheader_avg_score_overall"])
                # Check if there's any data to chart (sum of scores is > 0)
                if not df_avg.empty and df_avg["Average Score"].sum() > 0:
                    st.bar_chart(df_avg.set_index("Domain"))

                    # Add download buttons for Overall Domain Averages
                    col_download1, col_download2 = st.columns(2)
                    csv_buffer = io.StringIO()
                    df_avg.to_csv(csv_buffer, index=False)
                    col_download1.download_button(
                        label=strings["download_overall_avg_csv"],
                        data=csv_buffer.getvalue(),
                        file_name='overall_domain_averages.csv',
                        mime='text/csv',
                    )
                    excel_buffer = io.BytesIO()
                    df_avg.to_excel(excel_buffer, index=False)
                    col_download2.download_button(
                        label=strings["download_overall_avg_excel"],
                        data=excel_buffer.getvalue(),
                        file_name='overall_domain_averages.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    )

                else:
                     st.info(strings["info_no_numeric_scores_overall"])


                if not df_meta.empty:
                    st.subheader(strings["subheader_filter_analyze"])
                    # Use unique values from the dataframe for filters
                    col1, col2, col3, col4, col5 = st.columns(5) # Added column for teacher filter
                    school_filter = col1.selectbox(strings["filter_by_school"], ["All"] + sorted(df_meta["School"].dropna().unique().tolist()))
                    grade_filter = col2.selectbox(strings["filter_by_grade"], ["All"] + sorted(df_meta["Grade"].dropna().unique().tolist()))
                    subject_filter = col3.selectbox(strings["filter_by_subject"], ["All"] + sorted(df_meta["Subject"].dropna().unique().tolist()))

                    # Add date range filters
                    min_date = df_meta["Observation Date"].min() if pd.notnull(df_meta["Observation Date"].min()) else datetime.now().date()
                    max_date = df_meta["Observation Date"].max() if pd.notnull(df_meta["Observation Date"].max()) else datetime.now().date()

                    start_date = col4.date_input(strings["filter_start_date"], min_date)
                    end_date = col4.date_input(strings["filter_end_date"], max_date)

                    # Add teacher filter
                    teacher_list = ["All"] + sorted(df_meta["Teacher"].dropna().unique().tolist())
                    teacher_filter = col5.selectbox(strings["filter_teacher"], teacher_list)


                    # Filter metadata based on selection
                    filtered_meta_df = df_meta.copy()
                    if school_filter != "All":
                        filtered_meta_df = filtered_meta_df[filtered_meta_df["School"] == school_filter]
                    if grade_filter != "All":
                        filtered_meta_df = filtered_meta_df[filtered_meta_df["Grade"] == grade_filter]
                    if subject_filter != "All":
                        filtered_meta_df = filtered_meta_df[filtered_meta_df["Subject"] == subject_filter]
                    if teacher_filter != "All":
                         filtered_meta_df = filtered_meta_df[filtered_meta_df["Teacher"] == teacher_filter]


                    # Apply date filter - ensure 'Observation Date' is a datetime object or comparable
                    # Filter out rows where 'Observation Date' is None before comparison
                    filtered_meta_df = filtered_meta_df[filtered_meta_df['Observation Date'].notna()]
                    filtered_meta_df = filtered_meta_df[
                        (filtered_meta_df['Observation Date'] >= start_date) &
                        (filtered_meta_df['Observation Date'] <= end_date)
                    ]


                    # Filter the domain scores DataFrame based on the sheets in the filtered metadata
                    filtered_sheet_names = filtered_meta_df["Sheet"].tolist()
                    df_domain_scores_filtered = df_domain_scores_all[df_domain_scores_all['Sheet'].isin(filtered_sheet_names)].copy()

                    # Add download buttons for Filtered Observation Data and display filtered table
                    st.subheader("Filtered Observation Data") # Added subheader for filtered data table
                    if not filtered_meta_df.empty:
                         st.dataframe(filtered_meta_df) # Display filtered data

                         col_download3, col_download4 = st.columns(2)
                         csv_buffer_filtered_data = io.StringIO()
                         filtered_meta_df.to_csv(csv_buffer_filtered_data, index=False)
                         col_download3.download_button(
                             label=strings["download_filtered_data_csv"],
                             data=csv_buffer_filtered_data.getvalue(),
                             file_name='filtered_observation_data.csv',
                             mime='text/csv',
                         )
                         excel_buffer_filtered_data = io.BytesIO()
                         filtered_meta_df.to_excel(excel_buffer_filtered_data, index=False)
                         col_download4.download_button(
                             label=strings["download_filtered_data_excel"],
                             data=excel_buffer_filtered_data.getvalue(),
                             file_name='filtered_observation_data.xlsx',
                             mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         )
                    else:
                        st.info(strings["info_no_observation_data_filtered"])


                    # --- Teacher Performance Over Time ---
                    if teacher_filter != "All":
                        st.subheader(strings["subheader_teacher_performance"].format(teacher_filter))

                        if not df_domain_scores_filtered.empty:
                            # Calculate overall average for the filtered observations of this teacher
                            # Ensure 'Overall Score' column exists and has non-NA values
                            # Convert 'Overall Score' to numeric, coercing errors to NaN
                            df_domain_scores_filtered['Overall Score'] = pd.to_numeric(df_domain_scores_filtered['Overall Score'], errors='coerce')
                            teacher_overall_avg = df_domain_scores_filtered['Overall Score'].mean() if df_domain_scores_filtered['Overall Score'].notna().any() else None

                            st.subheader(strings["subheader_teacher_overall_avg"].format(teacher_filter))
                            if teacher_overall_avg is not None:
                                st.markdown(strings["overall_score_value"].format(teacher_overall_avg))
                            else:
                                st.markdown(strings["overall_score_na"])


                            # Prepare data for domain trend chart
                            # Need to melt the DataFrame to have 'Date', 'Domain', 'Score' columns
                            df_domain_scores_filtered_melted = df_domain_scores_filtered.melt(
                                id_vars=['Observation Date'],
                                value_vars=all_domains_list,
                                var_name='Domain',
                                value_name='Average Score'
                            )

                            # Drop rows where Average Score is None/NaN for charting
                            df_domain_scores_filtered_melted = df_domain_scores_filtered_melted.dropna(subset=['Average Score'])

                            if not df_domain_scores_filtered_melted.empty:
                                st.subheader(strings["subheader_teacher_domain_trend"].format(teacher_filter))

                                fig, ax = plt.subplots(figsize=(10, 6))
                                # Plot each domain's trend
                                for domain in all_domains_list:
                                    domain_data = df_domain_scores_filtered_melted[df_domain_scores_filtered_melted['Domain'] == domain]
                                    if not domain_data.empty:
                                         ax.plot(domain_data['Observation Date'], domain_data['Average Score'], marker='o', linestyle='-', label=domain)

                                ax.set_xlabel("Observation Date")
                                ax.set_ylabel("Average Score")
                                ax.set_title(strings["subheader_teacher_domain_trend"].format(teacher_filter))
                                ax.legend(title="Domains", bbox_to_anchor=(1.05, 1), loc='upper left')
                                plt.xticks(rotation=45, ha='right')
                                plt.tight_layout() # Adjust layout to prevent labels overlapping
                                st.pyplot(fig)

                            else:
                                st.info("No numeric domain scores found for the selected teacher within the applied filters to show trend.")


                        else:
                            st.info(strings["info_no_obs_for_teacher"])


                    else:
                        # Display overall filtered domain averages and observer distribution if "All" teachers selected
                        # Calculate and display filtered averages
                        # Only calculate mean if there are scores collected for that domain after filtering
                        # Use df_domain_scores_filtered here which is already filtered by date, school, grade, subject
                        filtered_avg_scores = {
                            domain: round(df_domain_scores_filtered[domain].mean(), 2) if not df_domain_scores_filtered[domain].isnull().all() else 0
                            for domain in all_domains_list
                        }

                        # Convert to DataFrame for charting, ensure all domains are present even if avg is 0
                        filtered_avg_data = [{"Domain": d, "Average Score": filtered_avg_scores.get(d, 0)} for d in all_domains_list]
                        df_filtered_avg = pd.DataFrame(filtered_avg_data)


                        st.subheader(strings["subheader_avg_score_filtered"])
                        # Check if there's any data after filtering (sum of scores is > 0)
                        if not df_filtered_avg.empty and df_filtered_avg["Average Score"].sum() > 0:
                            st.bar_chart(df_filtered_avg.set_index("Domain"))

                            # Add download buttons for Filtered Domain Averages
                            col_download5, col_download6 = st.columns(2)
                            csv_buffer_filtered_avg = io.StringIO()
                            df_filtered_avg.to_csv(csv_buffer_filtered_avg, index=False)
                            col_download5.download_button(
                                label=strings["download_filtered_avg_csv"],
                                data=csv_buffer_filtered_avg.getvalue(),
                                file_name='filtered_domain_averages.csv',
                                mime='text/csv',
                            )
                            excel_buffer_filtered_avg = io.BytesIO()
                            df_filtered_avg.to_excel(excel_buffer_filtered_avg, index=False)
                            col_download6.download_button(
                                label=strings["download_filtered_avg_excel"],
                                data=excel_buffer_filtered_avg.getvalue(),
                                file_name='filtered_domain_averages.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            )

                        else:
                            st.info(strings["info_no_numeric_scores_filtered"])


                        st.subheader(strings["subheader_observer_distribution"])
                        # Use the already filtered_meta_df for observer counts
                        if not filtered_meta_df.empty:
                            observer_counts = filtered_meta_df["Observer"].value_counts()
                            if not observer_counts.empty:
                                 fig, ax = plt.subplots()
                                 observer_counts.plot(kind='pie', autopct='%1.1f%%', ax=ax)
                                 ax.set_ylabel("") # Hide the default y-label for pie charts
                                 st.pyplot(fig)
                            else:
                                st.info(strings["info_no_observer_data_filtered"])
                        else:
                            st.info(strings["info_no_observation_data_filtered"])


        except Exception as e:
             st.error(f"Error loading or processing workbook for analytics: {e}") # More specific error message in analytics section

    else:
        st.warning(strings["warning_default_not_found"].format(DEFAULT_FILE)) # Use the same warning as input page
