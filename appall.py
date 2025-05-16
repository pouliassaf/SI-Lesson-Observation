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
from openpyxl.utils import get_column_letter # Import get_column_letter

# Import ReportLab modules for PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image # Import Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
import re # Import regex for cleaning HTML tags

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
    # More detailed recommendations based on performance levels
    "plan_very_weak_overall": "Overall performance is Very Weak. A comprehensive support plan is required. Focus on fundamental teaching practices such as classroom management, lesson planning, and basic instructional strategies. Seek guidance from your mentor teacher and school leadership.",
    "plan_weak_overall": "Overall performance is Weak. A support plan is recommended. Identify 1-2 key areas for improvement from the observation and work with your mentor teacher to develop targeted strategies. Consider observing experienced colleagues in these areas.",
    "plan_weak_domain": "Performance in **{}** is Weak. Focus on developing skills related to: {}. Suggested actions include: [Specific action 1], [Specific action 2].", # Domain Name, specific elements
    "steps_acceptable_overall": "Overall performance is Acceptable. Continue to build on your strengths. Identify one area for growth to refine your practice and enhance student learning.",
    "steps_good_overall": "Overall performance is Good. You demonstrate effective teaching practices. Explore opportunities to share your expertise with colleagues, perhaps through informal mentoring or presenting successful strategies.",
    "steps_good_domain": "Performance in **{}** is Good. You demonstrate strong skills in this area. Consider exploring advanced strategies related to: {}. Suggested actions include: [Specific advanced action 1], [Specific advanced action 2].", # Domain Name, specific elements
    "steps_excellent_overall": "Overall performance is Excellent. You are a role model for effective teaching. Consider leading professional development sessions or mentoring less experienced teachers.",
    "steps_excellent_domain": "Performance in **{}** is Excellent. Your practice in this area is exemplary. Continue to innovate and refine your practice, perhaps by researching and implementing cutting-edge strategies related to: {}.", # Domain Name, specific elements
    "no_specific_plan_needed": "Performance is at an acceptable level or above. No immediate support plan required based on this observation. Focus on continuous improvement based on your professional goals."
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
    # More detailed recommendations based on performance levels
    "plan_very_weak_overall": "الأداء الإجمالي ضعيف جداً. تتطلب خطة دعم شاملة. ركز على الممارسات التعليمية الأساسية مثل إدارة الصف، وتخطيط الدرس، والاستراتيجيات التعليمية الأساسية. اطلب التوجيه من معلمك الموجه وقيادة المدرسة.", # Guessed translation - Enhanced
    "plan_weak_overall": "الأداء الإجمالي ضعيف. يوصى بخطة دعم. حدد 1-2 من المجالات الرئيسية للتحسين من الملاحظة واعمل مع معلمك الموجه لتطوير استراتيجيات مستهدفة. فكر في ملاحظة الزملاء ذوي الخبرة في هذه المجالات.", # Guessed translation - Enhanced
    "plan_weak_domain": "الأداء في **{}** ضعيف. ركز على تطوير المهارات المتعلقة بـ: {}. الإجراءات المقترحة تشمل: [إجراء محدد 1]، [إجراء محدد 2].", # Guessed translation - Enhanced
    "steps_acceptable_overall": "الأداء الإجمالي مقبول. استمر في البناء على نقاط قوتك. حدد مجالًا واحدًا للنمو لتحسين ممارستك وتعزيز تعلم الطلاب.", # Guessed translation - Enhanced
    "steps_good_overall": "الأداء الإجمالي جيد. أنت تظهر ممارسات تعليمية فعالة. استكشف فرص مشاركة خبرتك مع الزملاء، ربما من خلال التوجيه غير الرسمي أو تقديم استراتيجيات ناجحة.", # Guessed translation - Enhanced
    "steps_good_domain": "الأداء في **{}** جيد. أنت تظهر مهارات قوية في هذا المجال. فكر في استكشاف استراتيجيات متقدمة تتعلق بـ: {}. الإجراءات المقترحة تشمل: [إجراء متقدم محدد 1]، [إجراء متقدم محدد 2].", # Guessed translation - Enhanced
    "steps_excellent_overall": "الأداء الإجمالي ممتاز. أنت نموذج يحتذى به في التدريس الفعال. فكر في قيادة جلسات التطوير المهني أو توجيه المعلمين الأقل خبرة.", # Guessed translation - Enhanced
    "steps_excellent_domain": "الأداء في **{}** ممتاز. ممارستك في هذا المجال نموذجية. استمر في الابتكار وتحسين ممارستك، ربما من خلال البحث وتطبيق استراتيجيات حديثة تتعلق بـ: {}.", # Guessed translation - Enhanced
    "no_specific_plan_needed": "الأداء عند مستوى مقبول أو أعلى. لا توجد خطة دعم فورية مطلوبة بناءً على هذه الملاحظة. ركز على التحسين المستمر بناءً على أهدافك المهنية." # Guessed translation - Enhanced
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
        elif numer
