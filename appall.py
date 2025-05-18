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
    "title_lesson_input": "أداة إدخال زيارة صفية أسبوعية",
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
    "steps_good_overall": "الأداء الإجمالي جيد. أنت تظهر ممارسات تعليمية فعالة. استكشف فرص مشاركة خبرتك مع الزملاء، ربما من خلال التوجيه غير الرسمي أو تقديم استراتيجيات ناجحة.", # Needs
