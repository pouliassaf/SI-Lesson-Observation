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
        elif numeric_score >= 4.5: # Fixed: Added colon
            return strings["perf_level_good"]
        elif numeric_score >= 3.5: # Fixed: Added colon
            return strings["perf_level_acceptable"]
        elif numeric_score >= 2.5: # Fixed: Added colon
            return strings["perf_level_weak"]
        else:
            return strings["perf_level_very_weak"]
    except (ValueError, TypeError):
        return strings["overall_score_na"] # Handle cases where score is not a valid number


# --- Define ReportLab Styles (DEFINED ONCE) ---
# Get default stylesheet
styles = getSampleStyleSheet()

# Add custom styles if they don't exist
if 'Heading1Centered' not in styles:
    styles.add(ParagraphStyle(name='Heading1Centered', alignment=1, fontSize=16, spaceAfter=14, bold=1))
if 'Heading2' not in styles:
    styles.add(ParagraphStyle(name='Heading2', fontSize=12, spaceAfter=10, bold=1))
if 'Normal' not in styles:
    styles.add(ParagraphStyle(name='Normal', fontSize=10, spaceAfter=6))
if 'RubricDescriptor' not in styles:
    styles.add(ParagraphStyle(name='RubricDescriptor', fontSize=9, spaceAfter=4, leftIndent=18)) # Indent descriptors
if 'RubricDomainHeading' not in styles:
    styles.add(ParagraphStyle(name='RubricDomainHeading', fontSize=11, spaceAfter=8, bold=1)) # Style for domain headings in PDF
if 'RubricElementRating' not in styles:
    styles.add(ParagraphStyle(name='RubricElementRating', fontSize=10, spaceAfter=4, leftIndent=10)) # Style for element rating in PDF


# --- Function to generate PDF ---
def generate_observation_pdf(data, feedback_content, strings, rubric_domains_structure):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)

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
            st.error(f"Could not add logo for {school_name}: {e}") # Changed warning to error for visibility
            # Optionally add a text placeholder if logo fails
            # story.append(Paragraph(f"[{school_name} Logo Placeholder]", styles['Normal']))
    else:
        st.error(f"Logo file not found for {school_name} at {logo_path}. Using text title.") # Changed warning to error for visibility
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
                     # Remove bold markdown and other potential HTML tags from descriptors for PDF
                     cleaned_desc_para = re.sub(r'<.*?>', '', description_en).replace('**', '')
                     # Split by newline for separate paragraphs in PDF
                     desc_paragraphs = cleaned_desc_para.split('\n')
                     for desc_para in desc_paragraphs:
                          if desc_para.strip():
                               story.append(Paragraph(desc_para, styles['RubricDescriptor']))
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

            selected_option = st.selectbox(strings["select_sheet_or_create"], ["Create new"] + lo_sheets)

            if selected_option == "Create new":
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

            domain_colors = ["#e6f2ff", "#fff2e6", "#e6ffe6", "#f9e6ff", "#ffe6e6", "#f0f0f5", "#e6f9ff", # Guessed translation
