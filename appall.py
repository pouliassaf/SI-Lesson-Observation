import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime, timedelta # Import timedelta for time calculation
import os
import statistics
import pandas as pd
import matplotlib.pyplot as plt # This requires matplotlib to be installed
import csv
import math # Import math for isnan check

# --- Text Strings for Localization ---
# You need to replace the placeholder Arabic strings with actual translations
en_strings = {
    "page_title": "Lesson Observation Tool",
    "sidebar_select_page": "Choose a page:",
    "page_lesson_input": "Lesson Input",
    "page_analytics": "Observation Analytics",
    "title_lesson_input": "Weekly Lesson Observation Input Tool",
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
    "label_obs_type": "Observation Type",
    "option_individual": "Individual",
    "option_joint": "Joint",
    "subheader_rubric_scores": "Rubric Scores",
    "expander_rubric_descriptors": "Rubric Descriptors",
    "info_no_descriptors": "No rubric descriptors available for this element.",
    "label_rating_for": "Rating for {}",
    "checkbox_send_feedback": "✉️ Send Feedback to Teacher",
    "button_save_observation": "💾 Save Observation",
    "warning_fill_essential": "Please fill in all basic information fields before saving.",
    "success_data_saved": "Observation data saved to {} in {}",
    "error_saving_workbook": "Error saving workbook:",
    "download_workbook": "📥 Download updated workbook",
    "feedback_subject": "Lesson Observation Feedback",
    "feedback_greeting": "Dear {},\n\nYour lesson observation from {} has been saved.\n\n",
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
    "title_analytics": "Observation Analytics Dashboard",
    "warning_no_lo_sheets_analytics": "No 'LO ' sheets found in the workbook for analytics.",
    "subheader_avg_score_overall": "Average Score per Domain (Across all observations)",
    "info_no_numeric_scores_overall": "No numeric scores found across all observations to calculate overall domain averages.",
    "subheader_data_summary": "Observation Data Summary",
    "subheader_filter_analyze": "Filter and Analyze",
    "filter_by_school": "Filter by School",
    "filter_by_grade": "Filter by Grade",
    "filter_by_subject": "Filter by Subject",
    "option_all": "All",
    "subheader_avg_score_filtered": "Average Score per Domain (Filtered)",
    "info_no_numeric_scores_filtered": "No observations matching the selected filters contain numeric scores for domain averages.",
    "subheader_observer_distribution": "Observer Distribution (Filtered)",
    "info_no_observer_data_filtered": "No observer data found for the selected filters.",
    "info_no_observation_data_filtered": "No observation data found for the selected filters.",
    "error_loading_analytics": "Error loading or processing workbook for analytics:",
    "overall_score_label": "Overall Score:", # Label for displaying overall score
    "overall_score_value": "**{:.2f}**", # Format for displaying overall score
    "overall_score_na": "**N/A**", # Display for no numeric scores
    "arabic_toggle_label": "عرض باللغة العربية (Display in Arabic)",
    "feedback_log_sheet_name": "Feedback Log",
    "feedback_log_header": ["Sheet", "Teacher", "Email", "Observer", "School", "Subject", "Date", "Summary"],
    "download_feedback_log_csv": "📥 Download Feedback Log (CSV)",
    "error_generating_log_csv": "Error generating feedback log CSV:",

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
    "page_title": "أداة ملاحظة الدرس", # Placeholder
    "sidebar_select_page": "اختر صفحة:", # Placeholder
    "page_lesson_input": "إدخال الدرس", # Placeholder
    "page_analytics": "تحليلات الملاحظة", # Placeholder
    "title_lesson_input": "أداة إدخال ملاحظة الدرس الأسبوعية", # Placeholder
    "info_default_workbook": "استخدام مصنف القالب الافتراضي:", # Placeholder
    "warning_default_not_found": "لم يتم العثور على مصنف القالب الافتراضي '{}'. يرجى تحميل مصنف.", # Placeholder
    "error_opening_default": "خطأ في فتح ملف القالب الافتراضي:", # Placeholder
    "success_lo_sheets_found": "تم العثور على {} أوراق LO في المصنف.", # Placeholder
    "select_sheet_or_create": "حدد ورقة LO موجودة أو أنشئ واحدة جديدة:", # Placeholder
    "option_create_new": "إنشاء جديد", # Placeholder
    "success_sheet_created": "تم إنشاء ورقة جديدة: {}", # Placeholder
    "error_template_not_found": "خطأ: لم يتم العثور على ورقة القالب 'LO 1' في المصنف! لا يمكن إنشاء ورقة جديدة.", # Placeholder
    "subheader_filling_data": "ملء البيانات لـ: {}", # Placeholder
    "label_observer_name": "اسم المراقب", # Placeholder
    "label_teacher_name": "اسم المعلم", # Placeholder
    "label_teacher_email": "البريد الإلكتروني للمعلم", # Placeholder
    "label_operator": "المشغل", # Placeholder
    "label_school_name": "اسم المدرسة", # Placeholder
    "label_grade": "الصف", # Placeholder
    "label_subject": "المادة", # Placeholder
    "label_gender": "الجنس", # Placeholder
    "label_students": "عدد الطلاب", # Placeholder
    "label_males": "عدد الذكور", # Placeholder
    "label_females": "عدد الإناث", # Placeholder
    "label_time_in": "وقت الدخول", # Placeholder
    "label_time_out": "وقت الخروج", # Placeholder
    "label_lesson_duration": "🕒 **مدة الدرس:** {} دقيقة — _{}_", # Placeholder
    "duration_full_lesson": "درس كامل", # Placeholder
    "duration_walkthrough": "جولة سريعة", # Placeholder
    "warning_calculate_duration": "يرجى إدخال وقت الدخول ووقت الخروج لحساب المدة.", # Placeholder
    "warning_could_not_calculate_duration": "تعذر حساب مدة الدرس:", # Placeholder
    "label_period": "الفترة", # Placeholder
    "label_obs_type": "نوع الملاحظة", # Placeholder
    "option_individual": "فردي", # Placeholder
    "option_joint": "مشترك", # Placeholder
    "subheader_rubric_scores": "درجات الدليل", # Placeholder
    "expander_rubric_descriptors": "واصفات الدليل", # Placeholder
    "info_no_descriptors": "لا توجد واصفات دليل متاحة لهذا العنصر.", # Placeholder
    "label_rating_for": "التقييم لـ {}", # Placeholder
    "checkbox_send_feedback": "✉️ إرسال ملاحظات إلى المعلم", # Placeholder
    "button_save_observation": "💾 حفظ الملاحظة", # Placeholder
    "warning_fill_essential": "يرجى ملء جميع حقول المعلومات الأساسية قبل الحفظ.", # Placeholder
    "success_data_saved": "تم حفظ بيانات الملاحظة في {} في {}", # Placeholder
    "error_saving_workbook": "خطأ في حفظ المصنف:", # Placeholder
    "download_workbook": "📥 تنزيل المصنف المحدث", # Placeholder
    "feedback_subject": "ملاحظات ملاحظة الدرس", # Placeholder
    "feedback_greeting": "عزيزي {},\n\nتم حفظ ملاحظة درسك من {}.\n\n", # Placeholder
    "feedback_observer": "المراقب: {}\n", # Placeholder
    "feedback_duration": "المدة: {}\n", # Placeholder
    "feedback_subject_fb": "المادة: {}\n", # Placeholder
    "feedback_school": "المدرسة: {}\n\n", # Placeholder
    "feedback_summary_header": "إليك ملخص لتقييماتك بناءً على الدليل:\n\n", # Placeholder
    "feedback_domain_header": "**{}: {}**\n", # Placeholder
    "feedback_element_rating": "- **{}:** التقييم **{}**\n", # Placeholder
    "feedback_descriptor_for_rating": "  *واصف للتقييم {}:* {}\n", # Placeholder
    "feedback_overall_score": "\n**متوسط الدرجة الإجمالي:** {:.2f}\n\n", # Placeholder
    "feedback_domain_average": "  *متوسط المجال:* {:.2f}\n", # Placeholder
    "feedback_performance_summary": "**ملخص الأداء:**\n", # Placeholder
    "feedback_overall_performance": "الإجمالي: {}\n", # Placeholder
    "feedback_domain_performance": "{}: {}\n", # Placeholder
    "feedback_support_plan_intro": "\n**خطة الدعم الموصى بها:**\n", # Placeholder
    "feedback_next_steps_intro": "\n**الخطوات التالية المقترحة:**\n", # Placeholder
    "feedback_closing": "\nبناءً على هذه التقييمات، يرجى مراجعة المصنف المحدث للحصول على ملاحظات تفصيلية ومجالات التطوير.\n\n", # Placeholder
    "feedback_regards": "مع التحيات,\nفريق قيادة المدرسة", # Placeholder
    "success_feedback_generated": "تم إنشاء الملاحظات (محاكاة):\n\n", # Placeholder
    "success_feedback_log_updated": "تم تحديث سجل الملاحظات في {}", # Placeholder
    "error_updating_log": "خطأ في تحديث سجل الملاحظات في المصنف:", # Placeholder
    "title_analytics": "لوحة تحليلات الملاحظة", # Placeholder
    "warning_no_lo_sheets_analytics": "لم يتم العثور على أوراق 'LO ' في المصنف للتحليلات.", # Placeholder
    "subheader_avg_score_overall": "متوسط الدرجة لكل مجال (عبر جميع الملاحظات)", # Placeholder
    "info_no_numeric_scores_overall": "لم يتم العثور على درجات رقمية عبر جميع الملاحظات لحساب متوسطات المجال الإجمالية.", # Placeholder
    "subheader_data_summary": "ملخص بيانات الملاحظة", # Placeholder
    "subheader_filter_analyze": "تصفية وتحليل", # Placeholder
    "filter_by_school": "تصفية حسب المدرسة", # Placeholder
    "filter_by_grade": "تصفية حسب الصف", # Placeholder
    "filter_by_subject": "تصفية حسب المادة", # Placeholder
    "option_all": "الكل", # Placeholder
    "subheader_avg_score_filtered": "متوسط الدرجة لكل مجال (مصفى)", # Placeholder
    "info_no_numeric_scores_filtered": "لا توجد ملاحظات مطابقة للمرشحات المحددة تحتوي على درجات رقمية لمتوسطات المجال.", # Placeholder
    "subheader_observer_distribution": "توزيع المراقبين (مصفى)", # Placeholder
    "info_no_observer_data_filtered": "لم يتم العثور على بيانات المراقب للمرشحات المحددة.", # Placeholder
    "info_no_observation_data_filtered": "لم يتم العثور على بيانات الملاحظة للمرشحات المحددة.", # Placeholder
    "error_loading_analytics": "خطأ في تحميل أو معالجة المصنف للتحليلات:", # Placeholder
    "overall_score_label": "النتيجة الإجمالية:", # Placeholder
    "overall_score_value": "**{:.2f}**", # Placeholder
    "overall_score_na": "**غير متوفر**", # Placeholder
    "arabic_toggle_label": "عرض باللغة العربية (Display in Arabic)", # Placeholder - Keep English part for clarity
    "feedback_log_sheet_name": "سجل الملاحظات", # Placeholder
    "feedback_log_header": ["الورقة", "المعلم", "البريد الإلكتروني", "المراقب", "المدرسة", "المادة", "التاريخ", "الملخص"], # Placeholder
    "download_feedback_log_csv": "📥 تنزيل سجل الملاحظات (CSV)", # Placeholder
    "error_generating_log_csv": "خطأ في إنشاء سجل الملاحظات CSV:", # Placeholder

    # Performance Level Descriptors (Arabic) - **Translate these**
    "perf_level_very_weak": "ضعيف جداً", # Placeholder
    "perf_level_weak": "ضعيف", # Placeholder
    "perf_level_acceptable": "مقبول", # Placeholder
    "perf_level_good": "جيد", # Placeholder
    "perf_level_excellent": "ممتاز", # Placeholder

    # Support Plan / Next Steps Text (Arabic) - **Translate and Customize these extensively**
    "plan_very_weak_overall": "الأداء الإجمالي ضعيف جداً. تتطلب خطة دعم شاملة، تركز على الممارسات التعليمية الأساسية عبر مجالات متعددة.", # Placeholder
    "plan_weak_overall": "الأداء الإجمالي ضعيف. يوصى بخطة دعم، تستهدف المجالات الرئيسية للتحسين المحددة في الملاحظة.", # Placeholder
    "plan_weak_domain": "الأداء في {} ضعيف. ركز على تطوير المهارات المتعلقة بـ: {}", # Placeholder
    "steps_acceptable_overall": "الأداء الإجمالي مقبول. استمر في البناء على نقاط القوة وركز على تحسين الممارسات في مجالات محددة.", # Placeholder
    "steps_good_overall": "الأداء الإجمالي جيد. استكشف فرص مشاركة أفضل الممارسات وتوجيه الزملاء.", # Placeholder
    "steps_good_domain": "الأداء في {} جيد. فكر في استراتيجيات متقدمة تتعلق بـ: {}", # Placeholder
    "steps_excellent_overall": "الأداء الإجمالي ممتاز. أنت نموذج يحتذى به في التدريس الفعال. فكر في قيادة التطوير المهني.", # Placeholder
    "steps_excellent_domain": "الأداء في {} ممتاز. استمر في الابتكار وتحسين ممارستك.", # Placeholder
    "no_specific_plan_needed": "الأداء عند مستوى مقبول أو أعلى. لا توجد خطة دعم فورية مطلوبة بناءً على هذه الملاحظة.", # Placeholder
}

# --- Function to get strings based on language toggle ---
def get_strings(arabic_mode):
    return ar_strings if arabic_mode else en_strings

# --- Function to determine performance level based on score ---
def get_performance_level(score, strings):
    if score is None:
        return strings["overall_score_na"] # Or a specific string for no score
    if score >= 5.5: # Example thresholds - Adjust as needed
        return strings["perf_level_excellent"]
    elif score >= 4.5:
        return strings["perf_level_good"]
    elif score >= 3.5:
        return strings["perf_level_acceptable"]
    elif score >= 2.5:
        return strings["perf_level_weak"]
    else:
        return strings["perf_level_very_weak"]

# --- Streamlit App Layout ---
# Add Arabic toggle early to affect language throughout the app
arabic_mode = st.sidebar.toggle(en_strings["arabic_toggle_label"], False)
strings = get_strings(arabic_mode)

# Set page config using the selected language
st.set_page_config(page_title=strings["page_title"], layout="wide")

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
                if not all([observer, teacher, school, grade, subject, students, males, females, time_in, time_out]):
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

                    # Save domain average scores to the sheet (Optional - choose columns)
                    # Example: Save to columns AB onwards
                    domain_avg_start_col = ord('AB') # ASCII value of 'AB'
                    for domain, avg_score in domain_avg_scores.items():
                         col_letter = chr(domain_avg_start_col + int(domain.split(" ")[-1]) - 1) # AB, AC, AD...
                         ws[f"{col_letter}15"] = f"{domain} Avg"
                         ws[f"{col_letter}16"] = avg_score if avg_score is not None else "N/A"


                    save_path = f"updated_{sheet_name}.xlsx"
                    try:
                        wb.save(save_path)
                        st.success(strings["success_data_saved"].format(sheet_name, save_path))
                        with open(save_path, "rb") as f:
                            st.download_button(strings["download_workbook"], f, file_name=save_path)
                        # os.remove(save_path) # Consider keeping the file for a bit or saving to a specific temp directory

                    except Exception as e:
                         st.error(strings["error_saving_workbook"].format(e))


                    # Generate and send feedback
                    if send_feedback and teacher_email:
                        feedback_content = strings["feedback_greeting"].format(teacher, datetime.now().strftime('%Y-%m-%d'))
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
                                datetime.now().strftime("%Y-%m-%d %H:%M"), feedback_content[:500] + ("..." if len(feedback_content) > 500 else "") # Truncate summary
                            ])

                            # Save the workbook again to include the log entry
                            wb.save(save_path)
                            st.success(strings["success_feedback_log_updated"].format(save_path))

                        except Exception as e:
                            st.error(strings["error_updating_log"].format(e))

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
            st.error(strings["error_loading_analytics"].format(e))


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
                        "Overall Score": ws["AA15"].value # Added Overall Score from saved data
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

                st.subheader(strings["subheader_avg_score_overall"])
                # Check if there's any data to chart (sum of scores is > 0)
                if not df_avg.empty and df_avg["Average Score"].sum() > 0:
                    st.bar_chart(df_avg.set_index("Domain"))
                else:
                     st.info(strings["info_no_numeric_scores_overall"])


                df_meta = pd.DataFrame(metadata)
                if not df_meta.empty:
                    st.subheader(strings["subheader_data_summary"])
                    st.dataframe(df_meta) # Show the full metadata table

                    st.subheader(strings["subheader_filter_analyze"])
                    # Use unique values from the dataframe for filters
                    col1, col2, col3 = st.columns(3)
                    school_filter = col1.selectbox(strings["filter_by_school"], ["All"] + sorted(df_meta["School"].dropna().unique().tolist()))
                    grade_filter = col2.selectbox(strings["filter_by_grade"], ["All"] + sorted(df_meta["Grade"].dropna().unique().tolist()))
                    subject_filter = col3.selectbox(strings["filter_by_subject"], ["All"] + sorted(df_meta["Subject"].dropna().unique().tolist()))

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
                                    if not math.isnan(rating): # Check if it's not NaN
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


                    st.subheader(strings["subheader_avg_score_filtered"])
                    # Check if there's any data after filtering (sum of scores is > 0)
                    if not df_filtered_avg.empty and df_filtered_avg["Average Score"].sum() > 0:
                        st.bar_chart(df_filtered_avg.set_index("Domain"))
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
             st.error(strings["error_loading_analytics"].format(e))

    else:
        st.warning(strings["warning_default_not_found"].format(DEFAULT_FILE)) # Use the same warning as input page

