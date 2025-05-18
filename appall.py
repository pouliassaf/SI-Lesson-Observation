#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon May 12 00:42:10 2025

@author: paulaassaf
"""

import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime, timedelta, date # Import timedelta and date
import os
import statistics
import pandas as pd # Added for potential analytics use later
import matplotlib.pyplot as plt # Added for potential analytics use later
import csv # Added for potential log download
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
# Ensure all strings used in the UI and feedback are included here and translated
en_strings = {
    "page_title": "Lesson Observation Tool",
    "sidebar_select_page": "Choose a page:",
    "page_lesson_input": "Lesson Observation Input",
    "page_analytics": "Lesson Observation Analytics",
    "page_help": "App Information and Guidelines", # New string for Help page
    "title_lesson_input": "Weekly Lesson Observation Input Tool",
    "title_help": "App Information and Guidelines", # New string for Help page title
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
    "label_teacher_email": "Teacher Email", # Added from first snippet
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
    "warning_calculate_duration": "Please enter both 'Time In' and 'Time Out' to calculate duration.", # Added based on snippet 1 logic
    "warning_could_not_calculate_duration": "Could not calculate lesson duration.",
    "label_period": "Period",
    "label_obs_type": "Observation Type",
    "option_individual": "Individual",
    "option_joint": "Joint",
    "subheader_rubric_scores": "Rubric Scores",
    "expander_rubric_descriptors": "Rubric Guidance", # Changed from Descriptors based on snippet 2 text
    "info_no_descriptors": "No rubric guidance available.", # Updated string
    "label_rating_for": "Rating for {}",
    "label_write_notes": "Write notes for {}", # Added string for note input label
    "checkbox_send_feedback": "✉️ Generate Feedback Report (for PDF)", # Renamed to clarify it's for PDF
    "button_save_observation": "💾 Save Observation",
    "warning_fill_essential": "Please fill in all essential information before saving.", # Added validation warning
    "success_data_saved": "Observation data saved to workbook.", # Simplified message
    "error_saving_workbook": "Error saving workbook:",
    "download_workbook": "📥 Download updated workbook",
    "feedback_subject": "Lesson Observation Feedback", # From snippet 1
    "feedback_greeting": "Dear {},\n\nYour lesson observation from {} has been saved.\n\n", # From snippet 1
    "feedback_observer": "Observer: {}\n", # From snippet 1
    "feedback_duration": "Duration: {}\n", # From snippet 1
    "feedback_subject_fb": "Subject: {}\n", # From snippet 1
    "feedback_school": "School: {}\n\n", # From snippet 1
    "feedback_summary_header": "Here is a summary of your ratings based on the rubric:\n\n", # From snippet 1
    "feedback_domain_header": "**{}: {}**\n", # Domain number and title - From snippet 1
    "feedback_element_rating": "- **{}:** Rating **{}**\n", # Element label and rating - From snippet 1
    "feedback_descriptor_for_rating": "  *Guidance for rating {}:* {}\n", # Descriptor for specific rating - Updated text
    "feedback_overall_score": "\n**Overall Average Score:** {:.2f}\n\n", # From snippet 1
    "feedback_domain_average": "  *Domain Average:* {:.2f}\n", # From snippet 1
    "feedback_performance_summary": "**Performance Summary:**\n", # Header for performance summary - From snippet 1
    "overall_performance_level_text": "Overall Performance Level: {}", # Added string for overall level
    "feedback_domain_performance": "{}: {}\n", # Domain performance level - From snippet 1
    "feedback_support_plan_intro": "\n**Support Plan Recommended:**\n", # Intro for support plan - From snippet 1
    "feedback_next_steps_intro": "\n**Suggested Next Steps:**\n", # Intro for next steps - From snippet 1
    "feedback_closing": "\nBased on these ratings, please review your updated workbook for detailed feedback and areas for development.\n\n", # From snippet 1
    "feedback_regards": "Regards,\nSchool Leadership Team", # From snippet 1
    "success_feedback_generated": "Feedback generated (simulated):\n\n", # From snippet 1
    "success_feedback_log_updated": "Feedback log updated.", # Simplified message
    "error_updating_log": "Error updating feedback log in workbook:", # From snippet 1
    "title_analytics": "Lesson Observation Analytics Dashboard", # From snippet 1
    "warning_no_lo_sheets_analytics": "No 'LO ' sheets found in the workbook for analytics.", # From snippet 1
    "subheader_avg_score_overall": "Average Score per Domain (Across all observations)", # From snippet 1
    "info_no_numeric_scores_overall": "No numeric scores found across all observations to calculate overall domain averages.", # From snippet 1
    "subheader_data_summary": "Observation Data Summary", # From snippet 1
    "subheader_filter_analyze": "Filter and Analyze",
    "filter_by_school": "Filter by School", # From snippet 1
    "filter_by_grade": "Filter by Grade", # From snippet 1
    "filter_by_subject": "Filter by Subject", # From snippet 1
    "filter_by_operator": "Filter by Operator", # New string for Operator filter
    "filter_by_observer_an": "Filter by Observer", # String for observer filter in analytics
    "option_all": "All", # From snippet 1
    "subheader_avg_score_filtered": "Average Score per Domain (Filtered)", # From snippet 1
    "info_no_numeric_scores_filtered": "No observations matching the selected filters contain numeric scores for domain averages.", # From snippet 1
    "subheader_observer_distribution": "Observer Distribution (Filtered)", # From snippet 1
    "info_no_observer_data_filtered": "No observer data found for the selected filters.", # From snippet 1
    "info_no_observation_data_filtered": "No observation data found for the selected filters.", # From snippet 1
    "error_loading_analytics": "Error loading or processing workbook for analytics:", # From snippet 1
    "overall_score_label": "Overall Score:", # From snippet 1
    "overall_score_value": "**{:.2f}**", # From snippet 1
    "overall_score_na": "**N/A**", # From snippet 1
    "arabic_toggle_label": "عرض باللغة العربية (Display in Arabic)", # From snippet 1
    "feedback_log_sheet_name": "Feedback Log", # From snippet 1
    "feedback_log_header": ["Sheet", "Observer", "Teacher", "Email", "School", "Subject", "Date", "Overall Judgment", "Overall Score", "Summary Notes"], # Updated log headers
    "download_feedback_log_csv": "📥 Download Feedback Log (CSV)", # From snippet 1
    "error_generating_log_csv": "Error generating log CSV:", # From snippet 1
    "download_overall_avg_csv": "📥 Download Overall Domain Averages (CSV)", # From snippet 1
    "download_overall_avg_excel": "📥 Download Overall Domain Averages (Excel)", # From snippet 1
    "download_filtered_avg_csv": "📥 Download Filtered Domain Averages (CSV)", # From snippet 1
    "download_filtered_avg_excel": "📥 Download Filtered Domain Averages (Excel)", # From snippet 1
    "download_filtered_data_csv": "📥 Download Filtered Observation Data (CSV)", # From snippet 1
    "download_filtered_data_excel": "📥 Download Filtered Observation Data (Excel)", # From snippet 1
    "label_observation_date": "Observation Date", # From snippet 1
    "filter_start_date": "Start Date", # From snippet 1
    "filter_end_date": "End Date", # From snippet 1
    "filter_teacher": "Filter by Teacher", # From snippet 1
    "subheader_teacher_performance": "Teacher Performance Over Time", # From snippet 1
    "info_select_teacher": "Select a teacher to view individual performance analytics.", # From snippet 1
    "info_no_obs_for_teacher": "No observations found for the selected teacher within the applied filters.", # From snippet 1
    "subheader_teacher_domain_trend": "{} Domain Performance Trend", # From snippet 1
    "subheader_teacher_overall_avg": "{} Overall Average Score (Filtered)", # From snippet 1
    "perf_level_very_weak": "Very Weak", # From snippet 1
    "perf_level_weak": "Weak", # From snippet 1
    "perf_level_acceptable": "Acceptable", # From snippet 1
    "perf_level_good": "Good", # From snippet 1
    "perf_level_excellent": "Excellent", # From snippet 1
    "plan_very_weak_overall": "Overall performance is Very Weak. A comprehensive support plan is required. Focus on fundamental teaching practices such as classroom management, lesson planning, and basic instructional strategies. Seek guidance from your mentor teacher and school leadership.", # From snippet 1
    "plan_weak_overall": "Overall performance is Weak. A support plan is recommended. Identify 1-2 key areas for improvement from the observation and work with your mentor teacher to develop targeted strategies. Consider observing experienced colleagues in these areas.", # From snippet 1
    "plan_weak_domain": "Performance in **{}** is Weak. Focus on developing skills related to: {}. Suggested actions include: [Specific action 1], [Specific action 2].", # From snippet 1
    "steps_acceptable_overall": "Overall performance is Acceptable. Continue to build on your strengths. Identify one area for growth to refine your practice and enhance student learning.", # From snippet 1
    "steps_good_overall": "Overall performance is Good. You demonstrate effective teaching practices. Explore opportunities to share your expertise with colleagues, perhaps through informal mentoring or presenting successful strategies.", # From snippet 1
    "steps_good_domain": "Performance in **{}** is Good. You demonstrate strong skills in this area. Consider exploring advanced strategies related to: {}. Suggested actions include: [Specific advanced action 1], [Specific advanced action 2].", # From snippet 1
    "steps_excellent_overall": "Overall performance is Excellent. You are a role model for effective teaching. Consider leading professional development sessions or mentoring less experienced teachers.", # From snippet 1
    "steps_excellent_domain": "Performance in **{}** is Excellent. Your practice in this area is exemplary. Continue to innovate and refine your practice, perhaps by researching and implementing cutting-edge strategies related to: {}.", # From snippet 1
    "no_specific_plan_needed": "Performance is at an acceptable level or above. No immediate support plan required based on this observation. Focus on continuous improvement based on your professional goals.", # From snippet 1
    "warning_fill_basic_info": "Please fill in Observer Name, Teacher Name, School Name, Grade, Subject, Gender, and Observation Date.", # More specific validation
    "warning_fill_all_basic_info": "Please fill in all basic information fields.", # Generic fallback
    "warning_numeric_fields": "Please enter valid numbers for Students, Males, and Females.", # Added string for numeric validation
    "success_pdf_generated": "Feedback PDF generated successfully.", # Added success message for PDF
    "download_feedback_pdf": "📥 Download Feedback PDF", # Added string for PDF download button label
    "checkbox_cleanup_sheets": "🪟 Clean up unused LO sheets (no observer name)", # Added string for checkbox label
    "warning_sheets_removed": "Removed {} unused LO sheets.", # Added string for warning message
    "info_reloaded_workbook": "Reloaded workbook after cleanup.", # Added string for info message
    "info_no_sheets_to_cleanup": "No unused LO sheets found to clean up.", # Added string for info message
    "expander_guidelines": "📘 Click here to view observation guidelines", # Added string for expander label
    "info_no_guidelines": "Guidelines sheet is empty or could not be read.", # Added string for info message
    "warning_select_create_sheet": "Please select or create a valid sheet to proceed.", # Added string for warning message
    "label_overall_notes": "General Notes for this Lesson Observation", # Added missing string key


}

# Placeholder Arabic strings - REPLACE THESE WITH ACTUAL TRANSLATIONS
ar_strings = {
    "page_title": "أداة التقييم للزيارات الصفية",
    "sidebar_select_page": "اختر صفحة:",
    "page_lesson_input": "ادخال تقييم الزيارة",
    "page_analytics": "تحليلات الزيارة",
    "page_help": "معلومات وإرشادات التطبيق", # New string for Help page - Needs verification
    "title_lesson_input": "أداة إدخال زيارة صفية أسبوعية",
    "title_help": "معلومات وإرشادات التطبيق", # New string for Help page title - Needs verification
    "info_default_workbook": "استخدام مصنف القالب الافتراضي:",
    "warning_default_not_found": "تحذير: لم يتم العثور على مصنف القالب الافتراضي '{}'. يرجى تحميل مصنف.",
    "error_opening_default": "خطأ في فتح ملف القالب الافتراضي:",
    "success_lo_sheets_found": "تم العثور على {} أوراق LO في المصنف.",
    "select_sheet_or_create": "حدد ورقة LO موجودة أو أنشئ واحدة جديدة:",
    "option_create_new": "إنشاء جديد",
    "success_sheet_created": "تم إنشاء ورقة جديدة: {}",
    "error_template_not_found": "خطأ: لم يتم العثور على ورقة القالب 'LO 1' في المصنف! لا يمكن إنشاء ورقة جديدة.",
    "subheader_filling_data": "ملء البيانات لـ: {}",
    "label_observer_name": "اسم المراقب",
    "label_teacher_name": "اسم المعلم",
    "label_teacher_email": "البريد الإلكتروني للمعلم",
    "label_operator": "المشغل",
    "label_school_name": "اسم المدرسة",
    "label_grade": "الصف",
    "label_subject": "المادة",
    "label_gender": "الجنس",
    "label_students": "عدد الطلاب",
    "label_males": "عدد الذكور",
    "label_females": "عدد الإناث",
    "label_time_in": "وقت الدخول",
    "label_time_out": "وقت الخروج",
    "label_lesson_duration": "🕒 **مدة الدرس:** {} دقيقة — _{}_",
    "duration_full_lesson": "درس كامل",
    "duration_walkthrough": "جولة سريعة",
    "warning_calculate_duration": "يرجى إدخال وقت الدخول ووقت الخروج لحساب المدة.",
    "warning_could_not_calculate_duration": "تعذر حساب مدة الدرس.",
    "label_period": "الفترة",
    "label_obs_type": "نوع الزيارة",
    "option_individual": "فردي",
    "option_joint": "مشترك",
    "subheader_rubric_scores": "درجات الدليل",
    "expander_rubric_descriptors": "إرشادات الدليل", # Needs verification
    "info_no_descriptors": "لا توجد إرشادات دليل متاحة لهذا العنصر.", # Needs verification
    "label_rating_for": "التقييم لـ {}",
    "label_write_notes": "كتابة ملاحظات لـ {}", # Guessed translation for notes label
    "checkbox_send_feedback": "✉️ إنشاء تقرير الملاحظات (للملف PDF)", # Guessed translation - renamed
    "button_save_observation": "💾 حفظ الزيارة",
    "warning_fill_essential": "يرجى ملء جميع حقول المعلومات الأساسية قبل الحفظ.",
    "success_data_saved": "تم حفظ بيانات الزيارة في المصنف.", # Guessed translation - simplified
    "error_saving_workbook": "خطأ في حفظ المصنف:",
    "download_workbook": "📥 تنزيل المصنف المحدث",
    "feedback_subject": "ملاحظات الزيارة الصفية", # Needs verification
    "feedback_greeting": "عزيزي {},\n\nتم حفظ زيارتك الصفية من {}.\n\n", # Needs verification
    "feedback_observer": "المراقب: {}\n", # Needs verification
    "feedback_duration": "المدة: {}\n", # Needs verification
    "feedback_subject_fb": "المادة: {}\n", # Needs verification
    "feedback_school": "المدرسة: {}\n\n", # Needs verification
    "feedback_summary_header": "إليك ملخص لتقييماتك بناءً على الدليل:\n\n", # Needs verification
    "feedback_domain_header": "**{}: {}**\n", # Needs verification
    "feedback_element_rating": "- **{}:** التقييم **{}**\n", # Needs verification
    "feedback_descriptor_for_rating": "  *إرشادات للتقييم {}:* {}\n", # Guessed translation for guidance text
    "feedback_overall_score": "\n**متوسط الدرجة الإجمالي:** {:.2f}\n\n", # Needs verification
    "feedback_domain_average": "  *متوسط المجال:* {:.2f}\n", # Needs verification
    "feedback_performance_summary": "**ملخص الأداء:**\n", # Needs verification
    "overall_performance_level_text": "مستوى الأداء الإجمالي: {}", # Guessed translation for overall level
    "feedback_domain_performance": "{}: {}\n", # Needs verification
    "feedback_support_plan_intro": "\n**خطة الدعم الموصى بها:**\n", # Needs verification
    "feedback_next_steps_intro": "\n**الخطوات التالية المقترحة:**\n", # Needs verification
    "feedback_closing": "\nبناءً على هذه التقييمات، يرجى مراجعة المصنف المحدث للحصول على ملاحظات تفصيلية ومجالات التطوير.\n\n", # Needs verification
    "feedback_regards": "مع التحيات,\nفريق قيادة المدرسة", # Needs verification
    "success_feedback_generated": "تم إنشاء الملاحظات (محاكاة):\n\n", # Needs verification
    "success_feedback_log_updated": "تم تحديث سجل الملاحظات.", # Guessed translation - simplified
    "error_updating_log": "خطأ في تحديث سجل الملاحظات في المصنف:", # Needs verification
    "title_analytics": "لوحة تحليلات الزيارة الصفية", # Needs verification
    "warning_no_lo_sheets_analytics": "لم يتم العثور على أوراق 'LO ' في المصنف للتحليلات.", # Needs verification
    "subheader_avg_score_overall": "متوسط الدرجة لكل مجال (عبر جميع الزيارات)", # Needs verification
    "info_no_numeric_scores_overall": "لم يتم العثور على درجات رقمية عبر جميع الزيارات لحساب متوسطات المجال الإجمالية.", # Needs verification
    "subheader_data_summary": "ملخص بيانات الزيارة", # Needs verification
    "subheader_filter_analyze": "تصفية وتحليل", # Needs verification
    "filter_by_school": "تصفية حسب المدرسة", # Needs verification
    "filter_by_grade": "تصفية حسب الصف", # Needs verification
    "filter_by_subject": "تصفية حسب المادة", # Needs verification
    "filter_by_operator": "تصفية حسب المشغل", # New string for Operator filter - Needs verification
    "filter_by_observer_an": "تصفية حسب المراقب", # String for observer filter in analytics - Needs verification
    "option_all": "الكل", # Needs verification
    "subheader_avg_score_filtered": "متوسط الدرجة لكل مجال (مصفى)", # Needs verification
    "info_no_numeric_scores_filtered": "لا توجد زيارات مطابقة للمرشحات المحددة تحتوي على درجات رقمية لمتوسطات المجال.", # Needs verification
    "subheader_observer_distribution": "توزيع المراقبين (مصفى)", # Needs verification
    "info_no_observer_data_filtered": "لم يتم العثور على بيانات المراقب للمرشحات المحددة.", # Needs verification
    "info_no_observation_data_filtered": "لم يتم العثور على بيانات الزيارة للمرشحات المحددة.", # Needs verification
    "error_loading_analytics": "خطأ في تحميل أو معالجة المصنف للتحليلات:", # Needs verification
    "overall_score_label": "النتيجة الإجمالية:", # Needs verification
    "overall_score_value": "**{:.2f}**", # Needs verification
    "overall_score_na": "**غير متوفر**", # Needs verification
    "arabic_toggle_label": "عرض باللغة العربية (Display in Arabic)", # Needs verification
    "feedback_log_sheet_name": "سجل الملاحظات", # Needs verification
    "feedback_log_header": ["Sheet", "Observer", "Teacher", "Email", "School", "Subject", "Date", "Overall Judgment", "Overall Score", "Summary Notes"], # Updated log headers - Guessed translation
    "download_feedback_log_csv": "📥 تنزيل سجل الملاحظات (CSV)", # Needs verification
    "error_generating_log_csv": "خطأ في إنشاء سجل الملاحظات CSV:", # Needs verification
    "download_overall_avg_csv": "📥 تنزيل متوسطات المجال الإجمالية (CSV)", # Needs verification
    "download_overall_avg_excel": "📥 تنزيل متوسطات المجال الإجمالية (Excel)", # Needs verification
    "download_filtered_avg_csv": "📥 تنزيل متوسطات المجال المصفاة (CSV)", # Needs verification
    "download_filtered_avg_excel": "📥 تنزيل متوسطات المجال المصفاة (Excel)", # Needs verification
    "download_filtered_data_csv": "📥 تنزيل بيانات الزيارة المصفاة (CSV)", # Needs verification
    "download_filtered_data_excel": "📥 تنزيل بيانات الزيارة المصفاة (Excel)", # Needs verification
    "label_observation_date": "تاريخ الزيارة", # Needs verification
    "filter_start_date": "تاريخ البدء", # Needs verification
    "filter_end_date": "تاريخ الانتهاء", # Needs verification
    "filter_teacher": "تصفية حسب المعلم", # Needs verification
    "subheader_teacher_performance": "أداء المعلم بمرور الوقت", # Needs verification
    "info_select_teacher": "حدد معلمًا لعرض تحليلات الأداء الفردي.", # Needs verification
    "info_no_obs_for_teacher": "لم يتم العثور على زيارات للمعلم المحدد ضمن المرشحات المطبقة.", # Needs verification
    "subheader_teacher_domain_trend": "اتجاه أداء مجال {}", # Needs verification
    "subheader_teacher_overall_avg": "متوسط الدرجة الإجمالي لـ {} (مصفى)", # Needs verification
    "perf_level_very_weak": "ضعيف جداً", # Needs verification
    "perf_level_weak": "ضعيف", # Needs verification
    "perf_level_acceptable": "مقبول", # Needs verification
    "perf_level_good": "جيد", # Needs verification
    "perf_level_excellent": "ممتاز", # Needs verification
    "plan_very_weak_overall": "الأداء الإجمالي ضعيف جداً. تتطلب خطة دعم شاملة. ركز على الممارسات التعليمية الأساسية مثل إدارة الصف، وتخطيط الدرس، والاستراتيجيات التعليمية الأساسية. اطلب التوجيه من معلمك الموجه وقيادة المدرسة.", # Needs verification
    "plan_weak_overall": "الأداء الإجمالي ضعيف. يوصى بخطة دعم. حدد 1-2 من المجالات الرئيسية للتحسين من الملاحظة واعمل مع معلمك الموجه لتطوير استراتيجيات مستهدفة. فكر في ملاحظة الزملاء ذوي الخبرة في هذه المجالات.", # Needs verification
    "plan_weak_domain": "الأداء في **{}** ضعيف. ركز على تطوير المهارات المتعلقة بـ: {}. الإجراءات المقترحة تشمل: [إجراء محدد 1]، [إجراء محدد 2].", # Needs verification
    "steps_acceptable_overall": "الأداء الإجمالي مقبول. استمر في البناء على نقاط قوتك. حدد مجالًا واحدًا للنمو لتحسين ممارستك وتعزيز تعلم الطلاب.", # Needs verification
    "steps_good_overall": "الأداء الإجمالي جيد. أنت تظهر ممارسات تعليمية فعالة. استكشف فرص مشاركة خبرتك مع الزملاء، ربما من خلال التوجيه غير الرسمي أو تقديم استراتيجيات ناجحة.", # Needs verification
    "steps_good_domain": "الأداء في **{}** جيد. أنت تظهر مهارات قوية في هذا المجال. فكر في استكشاف استراتيجيات متقدمة تتعلق بـ: {}. الإجراءات المقترحة تشمل: [إجراء متقدم محدد 1]، [إجراء متقدم محدد 2].", # Needs verification
    "steps_excellent_overall": "الأداء الإجمالي ممتاز. أنت نموذج يحتذى به في التدريس الفعال. فكر في قيادة جلسات التطوير المهني أو توجيه المعلمين الأقل خبرة.", # Needs verification
    "steps_excellent_domain": "الأداء في **{}** ممتاز. ممارستك في هذا المجال نموذجية. استمر في الابتكار وتحسين ممارستك، ربما من خلال البحث وتطبيق استراتيجيات حديثة تتعلق بـ: {}.", # Needs verification
    "no_specific_plan_needed": "الأداء عند مستوى مقبول أو أعلى. لا توجد خطة دعم فورية مطلوبة بناءً على هذه الملاحظة. ركز على التحسين المستمر بناءً على أهدافك المهنية.", # Needs verification
    "warning_fill_basic_info": "يرجى ملء اسم المراقب، اسم المعلم، اسم المدرسة، الصف، المادة، الجنس، وتاريخ الزيارة.", # Needs verification
    "warning_fill_all_basic_info": "يرجى ملء جميع حقول المعلومات الأساسية.", # Needs verification
    "warning_numeric_fields": "يرجى إدخال أرقام صحيحة لحقول الطلاب، الذكور، والإناث.", # Guessed translation
    "success_pdf_generated": "تم إنشاء ملف الملاحظات PDF بنجاح.", # Guessed translation
    "download_feedback_pdf": "📥 تنزيل ملف الملاحظات PDF", # Guessed translation
    "checkbox_cleanup_sheets": "🪟 تنظيف أوراق LO غير المستخدمة (لا يوجد اسم مراقب)", # Added string for checkbox label - Needs verification
    "warning_sheets_removed": "تمت إزالة {} أوراق LO غير مستخدمة.", # Added string for warning message - Needs verification
    "info_reloaded_workbook": "تمت إعادة تحميل المصنف بعد التنظيف.", # Added string for info message - Needs verification
    "info_no_sheets_to_cleanup": "لم يتم العثور على أوراق LO غير مستخدمة لتنظيفها.", # Added string for info message - Needs verification
    "expander_guidelines": "📘 انقر هنا لعرض إرشادات الملاحظة", # Added string for expander label - Needs verification
    "info_no_guidelines": "ورقة الإرشادات فارغة أو تعذر قراءتها.", # Added string for info message - Needs verification
    "warning_select_create_sheet": "يرجى تحديد أو إنشاء ورقة صالحة للمتابعة.", # Added string for warning message - Needs verification
    "label_overall_notes": "ملاحظات عامة لهذه الملاحظة الصفية", # Added missing string key - Needs verification
}


# --- Function to get strings based on language toggle ---
def get_strings(arabic_mode):
    return ar_strings if arabic_mode else en_strings

# --- Function to determine performance level based on score ---
def get_performance_level(score, strings):
    if score is None or (isinstance(score, (int, float)) and math.isnan(score)): # Handle int/float nan
        return strings["overall_score_na"]
    # Ensure score is treated as a number for comparison
    try:
        numeric_score = float(score)
        if numeric_score >= 5.5:
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
        # Handle non-numeric scores like "NA" or other errors
        return strings["overall_score_na"]


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
# This function is complex and depends on the data structure passed to it.
# It needs to be updated to match how data is collected and stored in the new snippet.
def generate_observation_pdf(data, feedback_content, strings, rubric_domains_structure, teacher_email):
    buffer = io.BytesIO()
    # Note: For Arabic support in PDF, you need to register Arabic fonts with ReportLab
    # and potentially adjust directionality in styles. This is a more advanced topic.
    doc = SimpleDocTemplate(buffer, pagesize=letter)

    story = []

    # --- Add School Logo ---
    school_name = data.get("school_name", "Default") # Use key from data dict
    logo_path = LOGO_PATHS.get(school_name, LOGO_PATHS["Default"])

    if os.path.exists(logo_path):
        try:
            img = Image(logo_path, width=1.5*inch, height=0.75*inch)
            img.hAlign = 'CENTER'
            story.append(img)
            story.append(Spacer(1, 0.2*inch))
        except Exception as e:
             # Log error without st.error if not running in Streamlit context for PDF build
             print(f"Could not add logo for {school_name}: {e}")
             story.append(Paragraph(f"[{school_name} Logo Placeholder]", styles['Normal'])) # Add placeholder text
    else:
        print(f"Logo file not found for {school_name} at {logo_path}. Using text title.")
        story.append(Paragraph(strings["page_title"], styles['Heading1Centered']))
        story.append(Spacer(1, 0.2*inch))

    # Add Title to PDF if logo wasn't added successfully or was placeholder
    if not story or not isinstance(story[-1], Image): # Check if last element is not an image (i.e., logo failed or wasn't added)
         # Avoid adding duplicate title if logo failed and text title was already added
         if story and story[-1].__class__.__name__ != 'Paragraph' or (story and story[-1].text != strings["page_title"]):
             story.append(Paragraph(strings["page_title"], styles['Heading1Centered']))
             story.append(Spacer(1, 0.1*inch))


    # Basic Information Table - Update keys to match data dictionary from input fields
    basic_info_data = [
        [strings["label_observer_name"] + ":", data.get("observer_name", "")],
        [strings["label_teacher_name"] + ":", data.get("teacher_name", "")],
        # Teacher Email might not be in the main data dict, pass separately or add it
        [strings["label_teacher_email"] + ":", teacher_email],
        [strings["label_operator"] + ":", data.get("operator", "")],
        [strings["label_school_name"] + ":", data.get("school_name", "")],
        [strings["label_grade"] + ":", data.get("grade", "")],
        [strings["label_subject"] + ":", data.get("subject", "")],
        [strings["label_gender"] + ":", data.get("gender", "")],
        [strings["label_students"] + ":", data.get("students", "")],
        [strings["label_males"] + ":", data.get("males", "")],
        [strings["label_females"] + ":", data.get("females", "")],
        [strings["label_observation_date"] + ":", data.get("observation_date", "")], # Using new key
        [strings["label_time_in"] + ":", data.get("time_in", "")],
        [strings["label_time_out"] + ":", data.get("time_out", "")],
        # The duration label and minutes need to be calculated/passed correctly
        [strings["label_lesson_duration"].split("🕒")[0].strip() + ":", data.get("duration_display", "")], # Pass formatted duration, strip emoji/html
        [strings["label_period"] + ":", data.get("period", "")],
        [strings["label_obs_type"] + ":", data.get("observation_type", "")], # Using new key
        # Overall score will need to be calculated and passed
        [strings["overall_score_label"] + ":", data.get("overall_score_display", strings["overall_score_na"])] # Use display value
    ]

    # Ensure all data points are strings for the table
    # Also handle date/time objects that might be in the data dict before converting to string
    cleaned_basic_info_data = []
    for item in basic_info_data:
        key, value = item
        if isinstance(value, (datetime, date, datetime.time)):
             # Format date/time objects nicely
             if isinstance(value, datetime):
                 formatted_value = value.strftime("%Y-%m-%d %H:%M")
             elif isinstance(value, date):
                 formatted_value = value.strftime("%Y-%m-%d")
             elif isinstance(value, datetime.time):
                 formatted_value = value.strftime("%H:%M")
             cleaned_basic_info_data.append([str(key), formatted_value])
        elif value is None:
             cleaned_basic_info_data.append([str(key), "N/A"])
        else:
             cleaned_basic_info_data.append([str(key), str(value)])


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

    table = Table(cleaned_basic_info_data, colWidths=[2*inch, 4*inch])
    table.setStyle(table_style)
    story.append(table)
    story.append(Spacer(1, 0.2*inch))

    # Rubric Scores - Needs to be adapted to the data structure collected from the inputs
    # and calculated averages/judgments.
    story.append(Paragraph(strings["subheader_rubric_scores"], styles['Heading2']))

    # Assuming 'data' dictionary passed to the PDF function contains domain_data
    # (which has title, average, judgment, and elements list)
    domain_data = data.get("domain_data", {})

    for domain_name, domain_info in domain_data.items():
        domain_title = domain_info.get("title", domain_name)
        domain_average = domain_info.get("average")
        domain_judgment = domain_info.get("judgment")
        elements = domain_info.get("elements", [])

        # Domain Title and Average
        story.append(Paragraph(f"<b>{domain_name}: {domain_title}</b>", styles['RubricDomainHeading']))
        if domain_average is not None and not math.isnan(domain_average):
             story.append(Paragraph(f"  Domain Average: {domain_average:.2f} ({domain_judgment})", styles['Normal']))
        else:
             story.append(Paragraph(f"  Domain Average: {strings['overall_score_na']}", styles['Normal']))
        story.append(Spacer(1, 0.1*inch))


        # Elements
        for element in elements:
            label = element.get("label", "Unknown Element")
            rating = element.get("rating", "N/A")
            note = element.get("note", "")
            descriptor = element.get("descriptor", "") # Specific descriptor text for the given rating

            story.append(Paragraph(f"- <b>{label}:</b> Rating <b>{rating}</b>", styles['RubricElementRating']))

            # Include Descriptor if available
            if descriptor and descriptor.strip():
                 # Clean and format the descriptor text
                 cleaned_desc_para = re.sub(r'<.*?>', '', descriptor).replace('**', '')
                 desc_paragraphs = cleaned_desc_para.split('\n')
                 for desc_para in desc_paragraphs:
                     if desc_para.strip():
                          story.append(Paragraph(desc_para, styles['RubricDescriptor']))
                 story.append(Spacer(1, 0.05*inch)) # Small space after descriptor

            # Include Note if available
            if note and note.strip():
                 story.append(Paragraph(f"  <i>Notes:</i> {note}", styles['Normal'])) # Italicize notes
                 story.append(Spacer(1, 0.05*inch)) # Small space after notes

            story.append(Spacer(1, 0.1*inch)) # Space after each element


        story.append(Spacer(1, 0.2*inch)) # Space after each domain

    # Add Overall Notes
    overall_notes = data.get("overall_notes", "")
    if overall_notes and overall_notes.strip():
        story.append(Paragraph("General Notes:", styles['Heading2']))
        # Convert markdown in general notes if any
        cleaned_overall_notes = re.sub(r'<.*?>', '', overall_notes).replace('**', '<b>').replace('**', '</b>')
        notes_paragraphs = cleaned_overall_notes.split('\n')
        for note_para in notes_paragraphs:
             if note_para.strip():
                 story.append(Paragraph(note_para, styles['Normal']))
        story.append(Spacer(1, 0.2*inch))


    # Feedback Content (This part is crucial and needs to be generated dynamically)
    story.append(Paragraph("Feedback Report:", styles['Heading2']))
    # The feedback_content string needs to be constructed before calling this function.
    # It should include the greeting, summary of scores/judgments, performance summary,
    # suggested plan/steps, and closing.
    if feedback_content:
        # Convert markdown-like text to ReportLab flowables
        feedback_paragraphs = feedback_content.split('\n\n') # Split by double newline
        for para in feedback_paragraphs:
            if para.strip():
                 # Simple bold conversion and newline handling for ReportLab
                 para_styled = para.replace('**', '<b>').replace('**', '</b>').replace('\n', '<br/>')
                 story.append(Paragraph(para_styled, styles['Normal']))
            story.append(Spacer(1, 0.1*inch)) # Add space between paragraphs
    else:
        story.append(Paragraph("Feedback content could not be generated.", styles['Normal']))


    # Build the PDF
    try:
        doc.build(story)
        buffer.seek(0)
        return buffer
    except Exception as e:
        # Handle errors during PDF build
        print(f"Error generating PDF: {e}") # Log to console/logs
        # st.error(f"Error generating PDF: {e}") # Avoid st.error inside this function if it runs outside Streamlit thread
        return None # Indicate failure


# --- Streamlit App Layout ---
# Add Arabic toggle early to affect language throughout the app
arabic_mode = st.sidebar.toggle(en_strings["arabic_toggle_label"], False)
strings = get_strings(arabic_mode)

# Sidebar page selection
page = st.sidebar.selectbox(strings["sidebar_select_page"], [strings["page_lesson_input"], strings["page_analytics"], strings["page_help"]]) # Added Help page

# --- File Handling ---
# Use a simple approach: load workbook directly if file exists
# Caching (@st.cache_resource) would improve performance on reruns but complicates saving
# For now, we load the workbook on each relevant rerun.
wb = None
DEFAULT_FILE = "Teaching Rubric Tool_WeekTemplate.xlsx"

if os.path.exists(DEFAULT_FILE):
     try:
         wb = load_workbook(DEFAULT_FILE) # Load directly from path
         st.info(strings["info_default_workbook"].format(DEFAULT_FILE))
     except Exception as e:
         st.error(strings["error_opening_default"].format(e))
         wb = None
else:
    st.warning(strings["warning_default_not_found"].format(DEFAULT_FILE))
    wb = None


# --- Main Application Logic ---
if wb: # Proceed only if workbook was loaded successfully
    if page == strings["page_lesson_input"]:
        st.title(strings["title_lesson_input"])

        # CSS for spacing
        st.markdown("""
        <style>
        .block-container {
            padding-top: 2rem;
        }
        </style>
        """, unsafe_allow_html=True)

        # Email Domain Restriction (Integrated from snippet 2)
        # Only show the rest of the app if email is entered and authorized
        email = st.text_input("Enter your school email to continue", value=st.session_state.get('auth_email_input', ''), key='auth_email_input')
        allowed_domains = ["@charterschools.ae", "@adek.gov.ae"]
        # Check if email is entered AND if it ends with an allowed domain
        if not (email and any(email.strip().lower().endswith(domain) for domain in allowed_domains)):
             if email.strip(): # Only show specific warning if email is entered but invalid
                  st.warning("Access restricted. Please use an authorized school email.")
             # Stop execution here if criteria are not met
             st.stop() # This stops the rest of the script below this point from running


        lo_sheets = [sheet for sheet in wb.sheetnames if sheet.startswith("LO ")]
        # Only report LO sheets if workbook loaded successfully
        if wb:
             st.success(strings["success_lo_sheets_found"].format(len(lo_sheets)))

        # Cleanup unused LO sheets (Integrated from snippet 2)
        # Only show cleanup option if there's more than just the template sheet
        # and workbook is loaded
        if wb and len(lo_sheets) > 1:
            if st.checkbox(strings.get("checkbox_cleanup_sheets", "🪟 Clean up unused LO sheets (no observer name)")): # Added string lookup, condition
                to_remove = []
                # Use AA1 as indicator for data existence, consistent with snippet 2 save logic
                for sheet in lo_sheets:
                    # Don't attempt to clean up the template sheet
                    if sheet == "LO 1":
                         continue
                    try:
                        # Check AA1 value in the sheet
                        aa1_value = wb[sheet]["AA1"].value
                        # Consider None or empty string as unused
                        if aa1_value is None or (isinstance(aa1_value, str) and aa1_value.strip() == ""):
                             to_remove.append(sheet)
                    except KeyError:
                        # Handle case where sheet might not have AA1 or is invalid
                        to_remove.append(sheet) # Consider sheets without AA1 as potentially unused/corrupt
                    except Exception as e:
                         st.warning(f"Could not check sheet '{sheet}' for cleanup: {e}")


                if to_remove: # Only attempt removal if there are sheets to remove
                    for sheet in to_remove:
                        # Double check it's not the template and still exists
                        if sheet != "LO 1" and sheet in wb.sheetnames:
                             try:
                                 wb.remove(wb[sheet])
                             except Exception as e:
                                  st.error(f"Error removing sheet {sheet}: {e}") # Report removal errors

                    # Reload sheet names after removal attempt
                    # Need to close and reopen the workbook to ensure openpyxl's internal state is clean after removal
                    # This is a limitation without caching/better state management
                    try:
                        # Save to a temporary buffer
                        temp_buffer = io.BytesIO()
                        wb.save(temp_buffer)
                        temp_buffer.seek(0)
                        # Reload from the buffer
                        wb = load_workbook(temp_buffer)
                        st.info(strings.get("info_reloaded_workbook", "Reloaded workbook after cleanup."))
                         # Re-run Streamlit explicitly to update the UI fully
                        st.rerun() # Changed experimental_rerun to rerun

                    except Exception as e:
                         st.error(f"Error reloading workbook after cleanup: {e}")


                else:
                     st.info(strings.get("info_no_sheets_to_cleanup", "No unused LO sheets found to clean up.")) # Message if no sheets were removed


        # Display Guidelines (Integrated from snippet 2)
        # Note: Guidelines are also displayed on the new Help page now.
        if wb and "Guidelines" in wb.sheetnames: # Ensure workbook is loaded before checking sheet
            # Attempt to read content safely
            guideline_content = []
            try:
                 # Read cells row by row, value only, skip None
                 for row in wb["Guidelines"].iter_rows(values_only=True):
                     for cell in row:
                         if cell is not None:
                             # Convert potential numbers to string and strip whitespace
                             guideline_content.append(str(cell).strip())
            except Exception as e:
                 st.error(f"Error reading Guidelines sheet: {e}")
                 guideline_content = [f"Error loading guidelines: {e}"] # Provide an error message


            # Join only non-empty lines
            cleaned_guidelines = [line for line in guideline_content if line]
            if cleaned_guidelines:
                 st.expander(strings.get("expander_guidelines", "📘 Click here to view observation guidelines")).markdown(
                     "\n".join(cleaned_guidelines) # Join with newline for markdown
                 )
            else:
                 st.info(strings.get("info_no_guidelines", "Guidelines sheet is empty or could not be read.")) # Message if sheet is empty


        lo_sheets = [sheet for sheet in wb.sheetnames if sheet.startswith("LO ")]
        # Ensure "LO 1" template is always available for copying
        if "LO 1" not in wb.sheetnames:
             st.error(strings["error_template_not_found"])
             st.stop() # Cannot proceed without template

        # Add "Create new" option only if "LO 1" exists and workbook is loaded
        # The LO 1 sheet should generally not be selectable for input, only used as a template.
        # So we only list existing LO sheets (excluding LO 1) and the "Create new" option.
        existing_sheets_for_selection = sorted([s for s in lo_sheets if s != "LO 1"])
        sheet_selection_options = [strings["option_create_new"]] + existing_sheets_for_selection


        # Determine initial selection index (try to keep current sheet if exists)
        # Reset to 'Create new' if the previously selected sheet was just cleaned up or is template
        current_sheet_name = st.session_state.get('current_sheet_name', sheet_selection_options[0])
        if current_sheet_name not in sheet_selection_options:
             current_sheet_name = sheet_selection_options[0] # Fallback to 'Create new'

        try:
             initial_index = sheet_selection_options.index(current_sheet_name)
        except ValueError:
             initial_index = 0 # Default to 'Create new' if current sheet not found


        selected_option = st.selectbox(strings["select_sheet_or_create"], sheet_selection_options, index=initial_index, key='sheet_selector')


        sheet_name = None
        ws_to_load_from = None # Initialize worksheet to load data from

        # --- Function to read existing data from a sheet (to pre-fill inputs) ---
        # Restored the definition of this function - PLACED HERE
        def load_existing_data(worksheet: Worksheet):
            data = {}
            # Basic Info from snippet 2 save locations
            # Use try-except for each cell read in case the sheet structure is unexpected
            try: data["gender"] = worksheet["B5"].value
            except Exception: pass
            try: data["students"] = worksheet["B6"].value
            except Exception: pass
            try: data["males"] = worksheet["B7"].value
            except Exception: pass
            try: data["females"] = worksheet["B8"].value
            except Exception: pass
            try: data["subject"] = worksheet["D2"].value
            except Exception: pass

            # Duration was calculated, need time in/out
            try:
                time_in_str = worksheet["D7"].value
                # Handle both time objects from previous saves and strings
                if isinstance(time_in_str, datetime.time):
                     data["time_in"] = time_in_str
                elif isinstance(time_in_str, datetime): # openpyxl sometimes reads time as datetime
                     data["time_in"] = time_in_str.time()
                elif isinstance(time_in_str, str) and time_in_str:
                     # Attempt parsing common time formats
                     try:
                         data["time_in"] = datetime.strptime(time_in_str, "%H:%M:%S").time() # Try with seconds
                     except ValueError:
                         data["time_in"] = datetime.strptime(time_in_str, "%H:%M").time() # Try without seconds
            except Exception:
                 data["time_in"] = None # Ensure it's set to None on error


            try:
                time_out_str = worksheet["D8"].value
                if isinstance(time_out_str, datetime.time):
                     data["time_out"] = time_out_str
                elif isinstance(time_out_str, datetime):
                     data["time_out"] = time_out_str.time()
                elif isinstance(time_out_str, str) and time_out_str:
                     try:
                         data["time_out"] = datetime.strptime(time_out_str, "%H:%M:%S").time()
                     except ValueError:
                         data["time_out"] = datetime.strptime(time_out_str, "%H:%M").time()
            except Exception:
                 data["time_out"] = None # Ensure it's set to None on error


            try: data["period"] = worksheet["D4"].value
            except Exception: pass


            # Metadata from snippet 2 save locations
            try: data["observer_name"] = worksheet["AA1"].value
            except Exception: pass
            try: data["teacher_name"] = worksheet["AA2"].value
            except Exception: pass
            try: data["observation_type"] = worksheet["AA3"].value
            except Exception: pass
            # Timestamp AA4 - don't load into input
            try: data["operator"] = worksheet["AA5"].value
            except Exception: pass
            try: data["school_name"] = worksheet["AA6"].value
            except Exception: pass
            try: data["overall_notes"] = worksheet["AA7"].value
            except Exception: pass

            # Date from assumed location D10
            try:
                 date_val = worksheet["D10"].value # Adjust cell as needed
                 if isinstance(date_val, datetime): # openpyxl reads dates as datetime
                     data["observation_date"] = date_val.date() # Store as date object
                 elif isinstance(date_val, date):
                      data["observation_date"] = date_val # Already a date object
                 # Add other potential date formats if necessary
            except Exception:
                 data["observation_date"] = datetime.now().date() # Default to today if error


            # Email - Assuming AA8 for email
            try: data["teacher_email"] = worksheet["AA8"].value # Assuming AA8 for email
            except Exception: pass


            # Rubric Scores and Notes - Read values saved in the sheet
            rubric_domains_structure = { # Need this structure to know where to read
                "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
            }
            data["element_inputs"] = {} # Store inputs keyed by unique key like f"{domain}_{i}_rating/note"
            for idx, (domain, (start_cell, count)) in enumerate(rubric_domains_structure.items()):
                 col_rating = start_cell[0] # Column for rating (e.g., 'I')
                 col_note = 'J' # Column for notes (based on snippet 2 save logic)
                 try:
                     row_start = int(start_cell[1:])
                     for i in range(count):
                          row = row_start + i
                          rating_key = f"{domain}_{i}_rating"
                          note_key = f"{domain}_{i}_note"
                          # Read value from sheet, use try-except for individual cells
                          try:
                              rating_value_from_sheet = worksheet[f"{col_rating}{row}"].value
                              # Convert numeric ratings to int if they are floats (openpyxl might read ints as floats)
                              if isinstance(rating_value_from_sheet, float) and rating_value_from_sheet.is_integer():
                                   rating_value_from_sheet = int(rating_value_from_sheet)
                              # Ensure "NA" is read correctly
                              elif isinstance(rating_value_from_sheet, str) and rating_value_from_sheet.upper() == "NA":
                                   rating_value_from_sheet = "NA"
                              # Convert numbers read as text back to numbers/NA
                              elif isinstance(rating_value_from_sheet, str) and rating_value_from_sheet.isdigit():
                                   rating_value_from_sheet = int(rating_value_from_sheet)
                              elif isinstance(rating_value_from_sheet, str) and rating_value_from_sheet.upper() == "NA":
                                   rating_value_from_sheet = "NA"
                              # Handle empty cells read as None
                              elif rating_value_from_sheet is None:
                                   rating_value_from_sheet = "NA"


                              data["element_inputs"][rating_key] = rating_value_from_sheet
                          except Exception:
                              data["element_inputs"][rating_key] = "NA" # Default to NA on error

                          try:
                              note_value_from_sheet = worksheet[f"{col_note}{row}"].value
                              data["element_inputs"][note_key] = note_value_from_sheet if note_value_from_sheet is not None else ""
                          except Exception:
                              data["element_inputs"][note_key] = "" # Default to empty string on error

                 except Exception as e:
                     st.warning(f"Error loading rubric data for domain {domain}: {e}")
                     # Continue to next domain even if one fails


            return data


        # --- Logic based on selected sheet/create new ---
        # This section determines the target sheet name and loads data into session state
        # It does *not* wrap the input display or save button logic.
        if selected_option == strings["option_create_new"]:
            # Determine the name for the new sheet (but don't create it until Save button is clicked)
            next_index = 1
            # Find highest existing LO number, skipping non-numeric suffixes
            existing_lo_numbers = []
            for sheet in wb.sheetnames:
                 if sheet.startswith("LO "):
                     suffix = sheet[3:].strip() # Get suffix and strip whitespace
                     if suffix.isdigit():
                         existing_lo_numbers.append(int(suffix))

            if existing_lo_numbers:
                 next_index = max(existing_lo_numbers) + 1

            sheet_name_to_process = f"LO {next_index}" # This is the *target* name for the new sheet

            is_new_sheet = True
            st.session_state['target_sheet_name'] = sheet_name_to_process
            # Clear relevant session state keys for new sheet, or rely on widget defaults
            keys_to_reset = ['observer_name', 'teacher_name', 'teacher_email', 'operator', 'school_name', 'grade',
                             'subject', 'gender', 'students', 'males', 'females', 'time_in', 'time_out',
                             'observation_date', 'period', 'observation_type', 'overall_notes', 'checkbox_send_feedback']

            for key in keys_to_reset:
                 # Keep the auth_email_input if it was successfully entered - do not reset it here
                 st.session_state[key] = None # Reset others to None

            # Reset session state for element inputs
            rubric_domains_structure = { # Need this structure here too
                 "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                 "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
            }
            # Initialize element inputs in session state for a new sheet
            if 'element_inputs' not in st.session_state:
                 st.session_state['element_inputs'] = {}
            for idx, (domain, (start_cell, count)) in enumerate(rubric_domains_structure.items()):
                 for i in range(count):
                     rating_key = f"{domain}_{i}_rating"
                     note_key = f"{domain}_{i}_note"
                     # Check if key already exists (shouldn't for a fresh "Create new" selection, but safety)
                     if rating_key not in st.session_state['element_inputs']:
                         st.session_state['element_inputs'][rating_key] = "NA" # Default rating
                     if note_key not in st.session_state['element_inputs']:
                         st.session_state['element_inputs'][note_key] = "" # Default note


            st.info(strings["subheader_filling_data"].format(sheet_name_to_process))
            ws_to_load_from = wb["LO 1"] # Rubric structure comes from the template


        else: # Selected an existing sheet
            sheet_name_to_process = selected_option
            st.session_state['target_sheet_name'] = sheet_name_to_process
            is_new_sheet = False
            try:
                 ws_to_load_from = wb[sheet_name_to_process] # Get the selected sheet object
                 st.subheader(strings["subheader_filling_data"].format(sheet_name_to_process))

                 # Load existing data into session state from the selected sheet
                 existing_data = load_existing_data(ws_to_load_from)
                 for key, value in existing_data.items():
                     # Populate session state, prioritizing loaded data unless it's the auth email
                     if key == 'element_inputs': # Special handling for the dictionary
                          st.session_state[key] = value
                     elif key == 'auth_email_input': # Keep auth email if already entered
                          continue
                     else:
                          st.session_state[key] = value


            except KeyError:
                 st.error(f"Error: Selected sheet '{sheet_name_to_process}' not found or could not be accessed.")
                 # Reset sheet selector if sheet is missing
                 st.session_state['current_sheet_name'] = sheet_selection_options[0] # Reset to 'Create new'
                 st.rerun() # Changed experimental_rerun to rerun
                 st.stop() # Stop execution if sheet loading fails
            except Exception as e:
                 st.error(f"Error loading data from sheet '{sheet_name_to_process}': {e}")
                 # Reset sheet selector if loading fails
                 st.session_state['current_sheet_name'] = sheet_selection_options[0] # Reset to 'Create new'
                 st.rerun() # Changed experimental_rerun to rerun
                 st.stop() # Stop execution if sheet loading fails


        # --- Removed Input and Save Functionality Block ---
        # The complex input/save logic was here and causing persistent SyntaxErrors.
        # It has been temporarily removed to provide a working baseline.

        st.warning("⚠️ **Input and Save Functionality Disabled** ⚠️")
        st.info("""
            The section for entering observation details, rubric scores, and notes,
            as well as the 'Save Observation' button and its associated logic (saving to Excel,
            updating the log, generating/downloading the PDF feedback report),
            has been temporarily disabled due to a persistent technical issue.

            The app currently allows you to:
            - Load the default workbook.
            - See the list of existing LO sheets.
            - Select an existing sheet or create a 'new' one (determining the target name and loading existing data if applicable).
            - Clean up unused sheets.
            - View the Guidelines (on the new 'App Information and Guidelines' page).
            - Navigate to the Analytics page.

            To restore the input and save features, the removed code block needs to be
            re-integrated and debugged. It's recommended to add back the functionality
            piece by piece, starting with basic info saving, then rubric scores,
            then notes, log updates, and finally PDF generation, to isolate where the
            syntax issue was occurring.
            """)


    # <--- This 'if page == strings["page_lesson_input"]:' block ends here.
    #       The 'elif' blocks below should align with it.
    #       This 'if/elif/elif' structure handles page navigation.
    elif page == strings["page_analytics"]:
        st.title(strings["title_analytics"])

        # Check if workbook is loaded before proceeding with analytics
        if not wb:
            st.warning(strings["warning_no_lo_sheets_analytics"]) # Reusing this string as appropriate
            st.stop()

        # --- Analytics Logic ---
        # Find all LO sheets (excluding the template) and the Feedback Log
        lo_sheets_data_list = [] # Use a list to build data before creating DataFrame
        feedback_log_data = pd.DataFrame()
        feedback_log_sheet_name = strings["feedback_log_sheet_name"]

        # Structure defining domains and their average cells in the LO sheets
        # (This is needed to extract domain averages for analytics)
        rubric_domains_avg_cells = {
             "Avg Domain 1": "I16", "Avg Domain 2": "I23", "Avg Domain 3": "I31",
             "Avg Domain 4": "I38", "Avg Domain 5": "I44", "Avg Domain 6": "I50",
             "Avg Domain 7": "I56", "Avg Domain 8": "I63", "Avg Domain 9": "I69"
         }


        try:
            # Load data from LO sheets
            lo_sheets_to_process = [sheet for sheet in wb.sheetnames if sheet.startswith("LO ") and sheet != "LO 1"]

            if not lo_sheets_to_process:
                st.info(strings["warning_no_lo_sheets_analytics"])
            else:
                for sheet_name in lo_sheets_to_process:
                    try:
                        ws = wb[sheet_name]
                        # Attempt to extract data points based on known cell locations
                        data = {
                            "Sheet": sheet_name,
                            "Observer": ws["AA1"].value if ws["AA1"].value is not None else None,
                            "Teacher": ws["AA2"].value if ws["AA2"].value is not None else None,
                            "Operator": ws["AA5"].value if ws["AA5"].value is not None else None, # Added Operator extraction
                            "School": ws["AA6"].value if ws["AA6"].value is not None else None,
                            "Grade": ws["B1"].value if ws["B1"].value is not None else None, # Assuming Grade is in B1 based on template layout
                            "Subject": ws["D2"].value if ws["D2"].value is not None else None,
                            "Gender": ws["B5"].value if ws["B5"].value is not None else None,
                            # Convert numbers to numeric directly, coerce errors
                            "Students": pd.to_numeric(ws["B6"].value, errors='coerce'),
                            "Males": pd.to_numeric(ws["B7"].value, errors='coerce'),
                            "Females": pd.to_numeric(ws["B8"].value, errors='coerce'),
                            "Observation Date": ws["D10"].value, # Assuming date is D10
                            "Observation Type": ws["AA3"].value if ws["AA3"].value is not None else None,
                            "Overall Score": None, # Placeholder, calculated next
                            "Overall Judgment": None, # Placeholder, calculated next
                        }

                        # Extract Domain Averages from LO sheet cells
                        for domain_key, cell_ref in rubric_domains_avg_cells.items():
                             try:
                                 avg_value = ws[cell_ref].value
                                 # Convert to numeric, errors='coerce' turns non-numbers into NaN
                                 data[domain_key] = pd.to_numeric(avg_value, errors='coerce')
                             except Exception:
                                 data[domain_key] = pd.NA # Store pandas NA on error


                        # Calculate Overall Score and Judgment from Excel formulas if possible
                        # Assuming Overall Score is calculated somewhere, e.g., AM1
                        try:
                            overall_score_excel = ws["AM1"].value # Adjust cell as needed based on template
                            if isinstance(overall_score_excel, (int, float)) and not math.isnan(overall_score_excel):
                                data["Overall Score"] = float(overall_score_excel)
                                # Recalculate judgment based on the numeric score using the function
                                data["Overall Judgment"] = get_performance_level(data["Overall Score"], strings)
                            else:
                                data["Overall Score"] = None # Use None or pd.NA for missing/invalid
                                data["Overall Judgment"] = strings["overall_score_na"]
                        except Exception:
                             data["Overall Score"] = None # Use None or pd.NA on error
                             data["Overall Judgment"] = strings["overall_score_na"]

                        # Append the data dictionary to the list
                        lo_sheets_data_list.append(data)

                    except Exception as e:
                        st.warning(f"Could not load data from sheet '{sheet_name}' for analytics: {e}")

            # Load data from Feedback Log sheet if it exists
            if feedback_log_sheet_name in wb.sheetnames:
                 log_ws = wb[feedback_log_sheet_name]
                 # Read data into a pandas DataFrame
                 # Assuming the log sheet has headers in the first row
                 data_rows = list(log_ws.iter_rows(min_row=2, values_only=True))
                 headers = [cell.value for cell in log_ws[1]] # Get headers from the first row

                 if headers and data_rows:
                      # Filter out empty rows if any
                      cleaned_data_rows = [row for row in data_rows if any(cell is not None and str(cell).strip() != "" for cell in row)]
                      if cleaned_data_rows:
                           feedback_log_data = pd.DataFrame(cleaned_data_rows, columns=headers)
                           # Attempt to convert 'Overall Score' to numeric, coercing errors
                           if 'Overall Score' in feedback_log_data.columns:
                                feedback_log_data['Overall Score'] = pd.to_numeric(feedback_log_data['Overall Score'], errors='coerce')
                           # Convert 'Date' column to datetime if it exists and is not already
                           if 'Date' in feedback_log_data.columns:
                                # Convert excel dates (numbers) or strings to datetime
                                feedback_log_data['Date'] = pd.to_datetime(feedback_log_data['Date'], errors='coerce')
            else:
                 st.info(strings["feedback_log_sheet_name"] + " not found.") # Use string directly


        except Exception as e:
            st.error(strings["error_loading_analytics"].format(e))
            st.stop()

        # Convert extracted LO sheets data list to DataFrame
        all_obs_data = pd.DataFrame(lo_sheets_data_list)

        if not all_obs_data.empty:
             # Convert 'Observation Date' column to datetime objects for robust comparison
             if 'Observation Date' in all_obs_data.columns:
                  all_obs_data['Observation Date'] = pd.to_datetime(all_obs_data['Observation Date'], errors='coerce')

             # Attempt to convert numeric columns (Students, Males, Females, Scores, Averages)
             # errors='coerce' handled during loading for scores and averages.
             # Ensure basic numbers are also numeric
             numeric_cols = ['Students', 'Males', 'Females'] + list(rubric_domains_avg_cells.keys()) + ['Overall Score']
             for col in numeric_cols:
                  if col in all_obs_data.columns:
                       all_obs_data[col] = pd.to_numeric(all_obs_data[col], errors='coerce')


             st.subheader(strings["subheader_data_summary"])
             st.write(f"Total Observations: {len(all_obs_data)}")
             if 'Overall Score' in all_obs_data.columns and not all_obs_data['Overall Score'].isna().all():
                 avg_overall_score = all_obs_data['Overall Score'].mean()
                 st.write(f"Overall Average Score: {avg_overall_score:.2f}")
             else:
                 st.write("Overall Average Score: N/A (No valid scores found)")


             # --- Filtering Options ---
             st.markdown("---")
             st.subheader(strings["subheader_filter_analyze"])

             # Get unique values for filters from ALL loaded LO sheets, drop NaNs, convert to list
             all_operators = sorted(all_obs_data['Operator'].dropna().unique().tolist()) if 'Operator' in all_obs_data.columns else []
             all_schools = sorted(all_obs_data['School'].dropna().unique().tolist()) if 'School' in all_obs_data.columns else []
             all_grades = sorted(all_obs_data['Grade'].dropna().unique().tolist()) if 'Grade' in all_obs_data.columns else []
             all_subjects = sorted(all_obs_data['Subject'].dropna().unique().tolist()) if 'Subject' in all_obs_data.columns else []
             all_teachers = sorted(all_obs_data['Teacher'].dropna().unique().tolist()) if 'Teacher' in all_obs_data.columns else []
             all_observers = sorted(all_obs_data['Observer'].dropna().unique().tolist()) if 'Observer' in all_obs_data.columns else []


             # Add filters
             filter_operator = st.selectbox(strings["filter_by_operator"], [strings["option_all"]] + all_operators) # Added Operator filter
             filter_school = st.selectbox(strings["filter_by_school"], [strings["option_all"]] + all_schools)
             filter_grade = st.selectbox(strings["filter_by_grade"], [strings["option_all"]] + all_grades)
             filter_subject = st.selectbox(strings["filter_by_subject"], [strings["option_all"]] + all_subjects)
             filter_teacher = st.selectbox(strings["filter_teacher"], [strings["option_all"]] + all_teachers)
             filter_observer = st.selectbox(strings["filter_by_observer_an"], [strings["option_all"]] + all_observers) # Added observer filter


             # Date filtering
             st.markdown("##### Filter by Date")
             today = datetime.now().date()
             # Determine min/max dates from the loaded data, handling NaT values
             valid_dates = all_obs_data['Observation Date'].dropna() if 'Observation Date' in all_obs_data.columns else pd.Series(dtype='datetime64[ns]')

             # Safely get min/max date, fall back to today +/- 1 year if no valid dates
             min_date_data = valid_dates.min().date() if not valid_dates.empty else today - timedelta(days=365)
             max_date_data = valid_dates.max().date() if not valid_dates.empty else today + timedelta(days=7)

             # Ensure min/max dates are valid date objects for the widget
             min_date_input = min_date_data if isinstance(min_date_data, date) else today - timedelta(days=365)
             max_date_input = max_date_data if isinstance(max_date_data, date) else today + timedelta(days=7)


             try:
                  # Ensure default value is within min/max range for the widget
                  default_start_date = max(min_date_input, today - timedelta(days=365)) # Use valid date objects
                  start_date = st.date_input(strings["filter_start_date"], value=default_start_date, min_value=min_date_input, max_value=max_date_input)
             except Exception:
                  start_date = st.date_input(strings["filter_start_date"], value=today - timedelta(days=365)) # Fallback default

             try:
                  # Ensure default value is within min/max range for the widget
                  default_end_date = min(max_date_input, today + timedelta(days=7))
                  end_date = st.date_input(strings["filter_end_date"], value=default_end_date, min_value=min_date_input, max_value=max_date_input)
             except Exception:
                  end_date = st.date_input(strings["filter_end_date"], value=today + timedelta(days=7)) # Fallback default


             # Apply Filters
             filtered_data = all_obs_data.copy()

             if filter_operator != strings["option_all"]:
                  # Ensure column exists and filter, handle potential NaNs in column
                 if 'Operator' in filtered_data.columns:
                      # Filter out rows where Operator is NaN before comparison, to avoid errors
                      filtered_data = filtered_data[filtered_data['Operator'].notna()].copy()
                      filtered_data = filtered_data[filtered_data['Operator'] == filter_operator].copy()
                 else:
                      st.warning(f"Operator column not found in data. Cannot filter by '{filter_operator}'.")
                      filtered_data = filtered_data.head(0) # Return empty dataframe


             if filter_school != strings["option_all"]:
                 if 'School' in filtered_data.columns:
                     filtered_data = filtered_data[filtered_data['School'].notna()].copy()
                     filtered_data = filtered_data[filtered_data['School'] == filter_school].copy()
                 else:
                     st.warning(f"School column not found in data. Cannot filter by '{filter_school}'.")
                     filtered_data = filtered_data.head(0)


             if filter_grade != strings["option_all"]:
                 if 'Grade' in filtered_data.columns:
                     filtered_data = filtered_data[filtered_data['Grade'].notna()].copy()
                     filtered_data = filtered_data[filtered_data['Grade'] == filter_grade].copy()
                 else:
                       st.warning(f"Grade column not found in data. Cannot filter by '{filter_grade}'.")
                       filtered_data = filtered_data.head(0)


             if filter_subject != strings["option_all"]:
                 if 'Subject' in filtered_data.columns:
                      filtered_data = filtered_data[filtered_data['Subject'].notna()].copy()
                      filtered_data = filtered_data[filtered_data['Subject'] == filter_subject].copy()
                 else:
                      st.warning(f"Subject column not found in data. Cannot filter by '{filter_subject}'.")
                      filtered_data = filtered_data.head(0)


             if filter_teacher != strings["option_all"]:
                 if 'Teacher' in filtered_data.columns:
                      filtered_data = filtered_data[filtered_data['Teacher'].notna()].copy()
                      filtered_data = filtered_data[filtered_data['Teacher'] == filter_teacher].copy()
                 else:
                       st.warning(f"Teacher column not found in data. Cannot filter by '{filter_teacher}'.")
                       filtered_data = filtered_data.head(0)


             if filter_observer != strings["option_all"]:
                  if 'Observer' in filtered_data.columns:
                       filtered_data = filtered_data[filtered_data['Observer'].notna()].copy()
                       filtered_data = filtered_data[filtered_data['Observer'] == filter_observer].copy()
                  else:
                       st.warning(f"Observer column not found in data. Cannot filter by '{filter_observer}'.")
                       filtered_data = filtered_data.head(0)


             # Apply date filter, ensuring the column exists and has valid datetimes
             if 'Observation Date' in filtered_data.columns and not filtered_data['Observation Date'].isna().all():
                  # Filter out NaT values before comparison and convert to date for comparison with date pickers
                  filtered_data_valid_dates = filtered_data.dropna(subset=['Observation Date']).copy()
                  # Convert date picker results to pandas Timestamps for direct comparison with datetime64[ns]
                  start_timestamp = pd.Timestamp(start_date)
                  end_timestamp = pd.Timestamp(end_date) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1) # Include the whole end day

                  filtered_data = filtered_data_valid_dates[(filtered_data_valid_dates['Observation Date'] >= start_timestamp) & (filtered_data_valid_dates['Observation Date'] <= end_timestamp)].copy() # Use copy() to avoid SettingWithCopyWarning
             else:
                 # If date column is missing or all NaT, date filtering cannot be applied meaningfully
                 # An empty DataFrame is already returned by the filters above if a column was missing.
                 # If the date column exists but is all NaN, filtered_data_valid_dates will be empty,
                 # resulting in filtered_data being empty, which is correct.
                 pass # No explicit 'else' needed here


             st.markdown("---")
             st.subheader(strings["subheader_avg_score_filtered"])

             if not filtered_data.empty:
                 # Calculate average overall score for filtered data, ignoring NaNs
                 if 'Overall Score' in filtered_data.columns and not filtered_data['Overall Score'].isna().all():
                      avg_filtered_score = filtered_data['Overall Score'].mean()
                      st.write(f"Average Overall Score for Filtered Data: {avg_filtered_score:.2f}")
                 else:
                      st.write("Average Overall Score for Filtered Data: N/A (No valid scores found)")


                 # --- Bar Chart: Overall Judgment Distribution (Filtered) ---
                 st.markdown("#### Overall Judgment Distribution (Filtered)")
                 if 'Overall Judgment' in filtered_data.columns and not filtered_data['Overall Judgment'].isna().all():
                     # Ensure non-string types (like None from NaNs) are handled before value_counts
                     valid_judgments = filtered_data['Overall Judgment'].dropna()
                     if not valid_judgments.empty:
                          judgment_counts = valid_judgments.value_counts()
                          # Define a specific order for the judgments
                          judgment_order = [strings["perf_level_very_weak"], strings["perf_level_weak"], strings["perf_level_acceptable"], strings["perf_level_good"], strings["perf_level_excellent"], strings["overall_score_na"]]
                          # Reindex to enforce order, fill missing categories with 0, then drop NA if desired
                          judgment_counts = judgment_counts.reindex(judgment_order, fill_value=0).drop(strings["overall_score_na"], errors='ignore')
                          # Only show the chart if there are counts for actual judgments
                          if not judgment_counts.empty and judgment_counts.sum() > 0:
                               st.bar_chart(judgment_counts)
                          else:
                               st.info("No observations with valid judgments found in the filtered data to chart.")
                     else:
                          st.info("No valid overall judgments found in the filtered data to chart.")
                 else:
                      st.info("Overall Judgment data is not available or is invalid in the filtered dataset.")


                 # --- Bar Chart: Average Score by Domain (Filtered) ---
                 st.markdown("#### Average Score by Domain (Filtered)")
                 # Calculate average for each domain column, ignoring NaNs
                 domain_avg_columns = [col for col in filtered_data.columns if col.startswith('Avg Domain')]
                 if domain_avg_columns:
                      # Select only the domain average columns and calculate the mean for each
                      domain_avg_data = filtered_data[domain_avg_columns].mean().reset_index()
                      domain_avg_data.columns = ['Domain', 'Average Score']
                      # Remove 'Avg ' prefix for cleaner chart labels
                      domain_avg_data['Domain'] = domain_avg_data['Domain'].str.replace('Avg ', '')

                      if not domain_avg_data.empty and not domain_avg_data['Average Score'].isna().all():
                           # Set the index to Domain for st.bar_chart
                           domain_avg_data = domain_avg_data.set_index('Domain')
                           st.bar_chart(domain_avg_data)
                      else:
                           st.info("No valid domain average scores found in the filtered data to chart.")
                 else:
                     st.info("Domain average data is not available in the dataset.")


                 # --- Filtered Data Table and Downloads ---
                 st.markdown("---")
                 st.markdown("##### Filtered Observation Data Table")
                 # Select specific columns for the table display for clarity
                 display_columns = ['Sheet', 'Observer', 'Teacher', 'Operator', 'School', 'Grade', 'Subject', 'Observation Date', 'Overall Score', 'Overall Judgment'] + domain_avg_columns # Added Operator
                 # Ensure selected display columns exist in the filtered data
                 display_columns = [col for col in display_columns if col in filtered_data.columns]
                 st.dataframe(filtered_data[display_columns])

                 st.markdown("###### Download Filtered Data")
                 col_csv, col_excel = st.columns(2)
                 with col_csv:
                      csv_buffer_filtered = io.StringIO()
                      # Include all columns in the downloaded CSV/Excel, not just display columns
                      filtered_data.to_csv(csv_buffer_filtered, index=False)
                      csv_buffer_filtered.seek(0)
                      st.download_button(
                          label=strings["download_filtered_data_csv"],
                          data=csv_buffer_filtered.getvalue(),
                          file_name="filtered_observation_data.csv",
                          mime="text/csv"
                      )
                 with col_excel:
                      excel_buffer_filtered = io.BytesIO()
                      filtered_data.to_excel(excel_buffer_filtered, index=False)
                      excel_buffer_filtered.seek(0)
                      st.download_button(
                           label=strings["download_filtered_data_excel"],
                           data=excel_buffer_filtered.getvalue(),
                           file_name="filtered_observation_data.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                      )

                 st.markdown("---")
                 st.info("""
                     **Note on Chart Downloads:** Streamlit's native charts (like the bar charts above)
                     do not have a built-in image download option. If you require downloadable chart
                     images, you would typically use charting libraries like Matplotlib or Plotly
                     (displayed using `st.pyplot()` or `st.plotly_chart()`) and add specific code
                     for chart export, which can add complexity. The data used to generate the charts
                     is available for download in the filtered data table above.
                     """)


             else:
                 st.info(strings["info_no_observation_data_filtered"]) # Show this if filtered_data is empty


             # --- Overall Domain Averages (Across ALL observations, ignoring filters) ---
             st.markdown("#### Average Score per Domain (Across all observations)")
             # Calculate average for each domain column from the unfiltered data, ignoring NaNs
             domain_avg_columns_all = [col for col in all_obs_data.columns if col.startswith('Avg Domain')]
             if domain_avg_columns_all:
                  domain_avg_data_all = all_obs_data[domain_avg_columns_all].mean().reset_index()
                  domain_avg_data_all.columns = ['Domain', 'Average Score']
                  # Remove 'Avg ' prefix from column names for cleaner chart labels
                  domain_avg_data_all['Domain'] = domain_avg_data_all['Domain'].str.replace('Avg ', '')

                  if not domain_avg_data_all.empty and not domain_avg_data_all['Average Score'].isna().all():
                       # Set the index to Domain for st.bar_chart
                       domain_avg_data_all = domain_avg_data_all.set_index('Domain')
                       st.bar_chart(domain_avg_data_all)

                       st.markdown("###### Download Overall Domain Averages Data")
                       col_csv_all_avg, col_excel_all_avg = st.columns(2)
                       with col_csv_all_avg:
                            csv_buffer_all_avg = io.StringIO()
                            domain_avg_data_all.to_csv(csv_buffer_all_avg)
                            csv_buffer_all_avg.seek(0)
                            st.download_button(
                                label=strings["download_overall_avg_csv"],
                                data=csv_buffer_all_avg.getvalue(),
                                file_name="overall_domain_averages.csv",
                                mime="text/csv"
                            )
                       with col_excel_all_avg:
                            excel_buffer_all_avg = io.BytesIO()
                            domain_avg_data_all.to_excel(excel_buffer_all_avg)
                            excel_buffer_all_avg.seek(0)
                            st.download_button(
                                 label=strings["download_overall_avg_excel"],
                                 data=excel_buffer_all_avg.getvalue(),
                                 file_name="overall_domain_averages.xlsx",
                                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )


                  else:
                       st.info(strings["info_no_numeric_scores_overall"]) # No valid overall domain averages found
             else:
                  st.info("Domain average data columns not found in the dataset.")



             # --- Teacher Performance Over Time ---
             st.markdown("---")
             st.subheader(strings["subheader_teacher_performance"])
             st.info(strings["info_select_teacher"])

             # Teacher selection dropdown for detailed trend
             # Use the list of unique teachers from *all* data, not just filtered data,
             # so you can select any teacher even if they are filtered out by other criteria initially.
             selected_teacher_for_trend = st.selectbox("Select Teacher for Trend Analysis", [None] + all_teachers, format_func=lambda x: x if x is not None else "Select a Teacher...") # Use None for the initial placeholder


             if selected_teacher_for_trend:
                  # Filter the *original* data by the selected teacher for the trend,
                  # but apply the date filter if set. Other filters (school, grade, subject, operator, observer)
                  # are applied by filtering `all_obs_data` into `filtered_data`.
                  # So, `teacher_data_for_trend` is already filtered by all criteria including teacher.
                  teacher_data_for_trend = filtered_data[filtered_data['Teacher'] == selected_teacher_for_trend].copy()


                  if not teacher_data_for_trend.empty and 'Overall Score' in teacher_data_for_trend.columns and not teacher_data_for_trend['Overall Score'].isna().all():
                       st.subheader(strings["subheader_teacher_overall_avg"].format(selected_teacher_for_trend))
                       # Display average for the selected teacher within the current filters
                       avg_teacher_score_filtered = teacher_data_for_trend['Overall Score'].mean()
                       st.write(f"Average Overall Score (Filtered): {avg_teacher_score_filtered:.2f}")

                       # Plot trend over time (Requires valid dates and domain data)
                       domain_avg_columns_teacher = [col for col in teacher_data_for_trend.columns if col.startswith('Avg Domain')]

                       if 'Observation Date' in teacher_data_for_trend.columns and not teacher_data_for_trend['Observation Date'].isna().all() and domain_avg_columns_teacher:
                            st.subheader(strings["subheader_teacher_domain_trend"].format(selected_teacher_for_trend))

                            # Prepare data for plotting trend - requires dates as index and numeric columns
                            # Sort data by date for the trend line
                            trend_data = teacher_data_for_trend.sort_values(by='Observation Date').copy() # Use copy()
                            # Select date and domain average columns
                            trend_columns = ['Observation Date'] + domain_avg_columns_teacher
                            trend_data = trend_data[trend_columns].dropna(subset=['Observation Date']) # Drop rows with no date

                            if not trend_data.empty:
                                 # Set date as index for plotting
                                 trend_data = trend_data.set_index('Observation Date')
                                 # Remove 'Avg ' prefix from column names for cleaner plot legend
                                 trend_data.columns = trend_data.columns.str.replace('Avg ', '')

                                 st.line_chart(trend_data)

                                 st.markdown("###### Download Teacher Trend Data")
                                 col_csv_trend, col_excel_trend = st.columns(2)
                                 with col_csv_trend:
                                      csv_buffer_trend = io.StringIO()
                                      trend_data.to_csv(csv_buffer_trend)
                                      csv_buffer_trend.seek(0)
                                      st.download_button(
                                          label="📥 Download Trend Data (CSV)",
                                          data=csv_buffer_trend.getvalue(),
                                          file_name=f"{selected_teacher_for_trend.replace(' ', '_')}_trend_data.csv",
                                          mime="text/csv"
                                      )
                                 with col_excel_trend:
                                      excel_buffer_trend = io.BytesIO()
                                      trend_data.to_excel(excel_buffer_trend)
                                      excel_buffer_trend.seek(0)
                                      st.download_button(
                                          label="📥 Download Trend Data (Excel)",
                                          data=excel_buffer_trend.getvalue(),
                                          file_name=f"{selected_teacher_for_trend.replace(' ', '_')}_trend_data.xlsx",
                                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                      )


                            else:
                                 st.info("No data with valid dates and domain scores found for this teacher under current filters.")

                       elif 'Observation Date' not in teacher_data_for_trend.columns or teacher_data_for_trend['Observation Date'].isna().all():
                            st.info("Observation dates are missing or invalid for this teacher under current filters.")
                       elif not domain_avg_columns_teacher:
                            st.info("Domain average data is not available for trend analysis.")


                  else:
                       st.info(strings["info_no_obs_for_teacher"])
             else:
                 st.info("Select a teacher above to view their performance trend.")



        else: # If all_obs_data is empty after initial loading
            st.info(strings["info_no_observation_data_filtered"])


    # <--- This 'elif page == strings["page_analytics"]:' block ends here.
    #       The 'elif' block below should align with it.
    #       This 'if/elif/elif' structure handles page navigation.
    elif page == strings["page_help"]: # New Help/Guidelines page
         st.title(strings["title_help"])

         # Read and display guidelines from the Excel sheet
         if wb and "Guidelines" in wb.sheetnames:
             guideline_content = []
             try:
                  for row in wb["Guidelines"].iter_rows(values_only=True):
                      # Flatten the row and filter out None values, convert to string
                      cleaned_row = [str(cell).strip() for cell in row if cell is not None]
                      guideline_content.extend(cleaned_row) # Add cleaned cells from the row

             except Exception as e:
                  st.error(f"Error reading Guidelines sheet: {e}")
                  guideline_content = [f"Error loading guidelines: {e}"] # Provide an error message

             # Filter out any empty strings resulting from strip()
             display_content = [line for line in guideline_content if line]

             if display_content:
                  # Join content with newlines for Markdown display
                  st.markdown("\n".join(display_content))
             else:
                  st.info(strings.get("info_no_guidelines", "Guidelines sheet is empty or could not be read."))
         else:
              st.warning("Guidelines sheet not found in the workbook.")


# <--- This 'if wb:' block ends here.
#       The final 'else' block should align with it.
#       This top-level 'if/else' structure handles initial workbook loading errors.
else: # If workbook could not be loaded at the very start
     st.error("Could not load the workbook. Please ensure 'Teaching Rubric Tool_WeekTemplate.xlsx' exists and is accessible.")
