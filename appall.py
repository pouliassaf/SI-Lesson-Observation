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

# --- Set Streamlit Page Config (MUST BE THE FIRST STREAMIT COMMAND) ---
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
page = st.sidebar.selectbox(strings["sidebar_select_page"], [strings["page_lesson_input"], strings["page_analytics"]])

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
                        st.experimental_rerun()

                    except Exception as e:
                         st.error(f"Error reloading workbook after cleanup: {e}")


                else:
                     st.info(strings.get("info_no_sheets_to_cleanup", "No unused LO sheets found to clean up.")) # Message if no sheets were removed


        # Display Guidelines (Integrated from snippet 2)
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
        sheet_selection_options = [strings["option_create_new"]] + sorted([s for s in lo_sheets if s != "LO 1"]) # Sort existing sheets alphabetically, exclude template
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
        ws = None # Initialize worksheet variable

        # Function to read existing data from a sheet (to pre-fill inputs)
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
        if selected_option == strings["option_create_new"]:
            # Determine the name for the new sheet
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

            sheet_name_to_process = f"LO {next_index}"

            # Signal that this is a new sheet, inputs should be empty/default
            is_new_sheet = True
            st.session_state['target_sheet_name'] = sheet_name_to_process
            # Clear relevant session state keys for new sheet, or rely on widget defaults
            # Explicitly setting None might be safer if widgets don't reset perfectly
            # Use a set of keys to reset for cleaner code
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
            for idx, (domain, (start_cell, count)) in enumerate(rubric_domains_structure.items()):
                 for i in range(count):
                     rating_key = f"{domain}_{i}_rating"
                     note_key = f"{domain}_{i}_note"
                     st.session_state[rating_key] = "NA" # Default rating for new sheet
                     st.session_state[note_key] = "" # Default note for new sheet


            st.info(strings["subheader_filling_data"].format(sheet_name_to_process))
            ws_to_load_from = wb["LO 1"] # Load rubric structure/descriptors from template
            st.session_state['current_sheet_name'] = sheet_name_to_process # Update session state here


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
                     # Populate session state only if the key doesn't exist or is None,
                     # to avoid overwriting user changes on rerun after load.
                     # Also handle specific cases like time/date where None might be a valid load value.
                     # Always load if the value from the sheet is not None/empty string (for basic info)
                     # or always load for element inputs as they represent the sheet state.
                     if key == 'element_inputs': # Special handling for the dictionary
                          st.session_state[key] = value
                     elif key in ['time_in', 'time_out', 'observation_date']: # Always load if not None for time/date
                          st.session_state[key] = value
                     elif key in ['students', 'males', 'females']: # Load numbers as strings
                          st.session_state[key] = str(value) if value is not None else ''
                     # Only load if the value from the sheet is not None and is different from current session state
                     # This prevents overwriting user changes if they quickly switch sheets
                     # But for robustness, let's prioritize the data from the loaded sheet when selecting a sheet
                     # except for the email which is auth
                     elif key != 'auth_email_input':
                          st.session_state[key] = value

                 st.session_state['current_sheet_name'] = selected_option # Update session state here

            except KeyError:
                 st.error(f"Error: Selected sheet '{sheet_name_to_process}' not found or could not be accessed.")
                 # Reset sheet selector if sheet is missing
                 st.session_state['current_sheet_name'] = sheet_selection_options[0] # Reset to 'Create new'
                 st.experimental_rerun() # Rerun to show the corrected state
                 st.stop() # Stop execution if sheet loading fails
            except Exception as e:
                 st.error(f"Error loading data from sheet '{sheet_name_to_process}': {e}")
                 # Reset sheet selector if loading fails
                 st.session_state['current_sheet_name'] = sheet_selection_options[0] # Reset to 'Create new'
                 st.experimental_rerun() # Rerun to show the corrected state
                 st.stop() # Stop execution if sheet loading fails


        # --- Inputs Section (Only display if a sheet is selected or created) ---
        # Proceed with inputs only if a target sheet name is determined AND workbook is loaded
        if wb and st.session_state.get('target_sheet_name') and st.session_state.get('target_sheet_name') != strings["option_create_new"]: # Ensure we have a valid target sheet name and workbook

             # --- Basic Information Inputs ---
             # Inputs now read/write directly to st.session_state based on their keys
             # Provide default values if session state keys are still None (e.g., after reset for new sheet)
             observer = st.text_input(strings["label_observer_name"], value=st.session_state.get('observer_name', ''), key='observer_name')
             teacher = st.text_input(strings["label_teacher_name"], value=st.session_state.get('teacher_name', ''), key='teacher_name')
             teacher_email = st.text_input(strings["label_teacher_email"], value=st.session_state.get('teacher_email', ''), key='teacher_email') # Added email
             operator_options = sorted(["Taaleem", "Al Dar", "New Century Education", "Bloom"])
             initial_operator_index = operator_options.index(st.session_state.get('operator', operator_options[0])) if st.session_state.get('operator') in operator_options else 0
             operator = st.selectbox(strings["label_operator"], operator_options, index=initial_operator_index, key='operator')


             school_options = { # Hardcoded - consider reading from Excel
                 "New Century Education": ["Al Bayan School", "Al Bayraq School", "Al Dhaher School", "Al Hosoon School", "Al Mutanabi School", "Al Nahdha School", "Jern Yafoor School", "Maryam Bint Omran School"],
                 "Taaleem": ["Al Ahad Charter School", "Al Azm Charter School", "Al Riyadh Charter School", "Al Majd Charter School", "Al Qeyam Charter School", "Al Nayfa Charter Kindergarten", "Al Salam Charter School", "Al Walaa Charter Kindergarten", "Al Forsan Charter Kindergarten", "Al Wafaa Charter Kindergarten", "Al Watan Charter School"],
                 "Al Dar": ["Al Ghad Charter School", "Al Mushrif Charter Kindergarten", "Al Danah Charter School", "Al Rayaheen Charter School", "Al Rayana Charter School", "Al Qurm Charter School", "Mubarak Bin Mohammed Charter School (Cycle 2 & 3)"], # Corrected ]
                 "Bloom": ["Al Ain Charter School", "Al Dana Charter School", "Al Ghadeer Charter School", "Al Hili Charter School", "Al Manhal Charter School", "Al Qattara Charter School", "Al Towayya Charter School", "Jabel Hafeet Charter School"]
             }
             school_list = sorted(school_options.get(operator, []))
             # Handle index safely, default to 0 if current operator has no schools or session state value is not in the current list
             initial_school_index = 0
             if st.session_state.get('school_name') in school_list:
                 initial_school_index = school_list.index(st.session_state.get('school_name'))
             school = st.selectbox(strings["label_school_name"], school_list, index=initial_school_index, key='school_name')


             grade_options = [f"Grade {i}" for i in range(1, 13)] + ["K1", "K2"]
             initial_grade_index = 0
             if st.session_state.get('grade') in grade_options:
                 initial_grade_index = grade_options.index(st.session_state.get('grade'))
             grade = st.selectbox(strings["label_grade"], grade_options, index=initial_grade_index, key='grade')

             subject_options = ["Math", "English", "Arabic", "Science", "Islamic", "Social Studies"] # Hardcoded - consider reading from Excel
             initial_subject_index = 0
             if st.session_state.get('subject') in subject_options:
                 initial_subject_index = subject_options.index(st.session_state.get('subject'))
             subject = st.selectbox(strings["label_subject"], subject_options, index=initial_subject_index, key='subject')

             gender_options = ["Male", "Female", "Mixed"]
             initial_gender_index = 0
             if st.session_state.get('gender') in gender_options:
                 initial_gender_index = gender_options.index(st.session_state.get('gender'))
             gender = st.selectbox(strings["label_gender"], gender_options, index=initial_gender_index, key='gender')


             students = st.text_input(strings["label_students"], value=st.session_state.get('students', ''), key='students')
             males = st.text_input(strings["label_males"], value=st.session_state.get('males', ''), key='males')
             females = st.text_input(strings["label_females"], value=st.session_state.get('females', ''), key='females')

             # Time inputs - Need to handle datetime.time objects in session state
             # time_input returns datetime.time or None
             time_in = st.time_input(strings["label_time_in"], value=st.session_state.get('time_in'), key='time_in')
             time_out = st.time_input(strings["label_time_out"], value=st.session_state.get('time_out'), key='time_out')

             # Date input - Need to handle date objects in session state
             # date_input returns datetime.date or None
             # Provide a default if session state is empty
             default_date_value = st.session_state.get('observation_date', datetime.now().date())
             observation_date = st.date_input(strings["label_observation_date"], value=default_date_value, key='observation_date')


             # Calculate and display lesson duration (Integrated from snippet 1/2)
             lesson_duration = None
             duration_label = "N/A"
             minutes = 0
             # Calculate duration if both time inputs have values in session state
             time_in_ss = st.session_state.get('time_in')
             time_out_ss = st.session_state.get('time_out')

             try:
                 if time_in_ss is not None and time_out_ss is not None:
                     dummy_date = date.today()
                     start_dt = datetime.combine(dummy_date, time_in_ss)
                     end_dt = datetime.combine(dummy_date, time_out_ss)

                     if end_dt < start_dt:
                         end_dt += timedelta(days=1)

                     lesson_duration = end_dt - start_dt
                     minutes = round(lesson_duration.total_seconds() / 60)
                     duration_label = strings["duration_full_lesson"] if minutes >= 40 else strings["duration_walkthrough"]
                     st.markdown(strings["label_lesson_duration"].format(minutes, duration_label))
                 else:
                     st.warning(strings["warning_calculate_duration"])
             except Exception as e:
                 st.warning(strings["warning_could_not_calculate_duration"].format(e))


             period_options = [f"Period {i}" for i in range(1, 9)]
             initial_period_index = 0
             if st.session_state.get('period') in period_options:
                 initial_period_index = period_options.index(st.session_state.get('period'))
             period = st.selectbox(strings["label_period"], period_options, index=initial_period_index, key='period')

             obs_type_options = [strings["option_individual"], strings["option_joint"]]
             initial_obstype_index = 0
             if st.session_state.get('observation_type') in obs_type_options:
                 initial_obstype_index = obs_type_options.index(st.session_state.get('observation_type'))
             obs_type = st.selectbox(strings["label_obs_type"], obs_type_options, index=initial_obstype_index, key='observation_type')


             # --- Rubric Scores Input (Integrated and Localized) ---
             st.markdown("---")
             st.subheader(strings["subheader_rubric_scores"])

             # Structure defining domains, start cells, and element counts
             rubric_domains_structure = {
                 "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                 "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
             }

             # Dictionary to store element labels, notes, and descriptors from template for PDF/Feedback
             rubric_template_data = {}

             domain_colors = ["#e6f2ff", "#fff2e6", "#e6ffe6", "#f9e6ff", "#ffe6e6", "#f0f0f5", "#e6f9ff", "#f2ffe6", "#ffe6f2"]

             try:
                 # Assuming "LO 1" sheet contains the rubric details
                 rubric_template_ws = wb["LO 1"]

                 for idx, (domain_name, (start_cell, count)) in enumerate(rubric_domains_structure.items()):
                     background = domain_colors[idx % len(domain_colors)]

                     # Read Domain Title from template
                     try:
                         domain_row_template = int(start_cell[1:]) # Row for Domain Title in template
                         domain_title = rubric_template_ws[f"A{domain_row_template}"].value or domain_name
                     except Exception:
                         domain_title = domain_name # Fallback if reading title fails

                     # Store domain title and prepare elements list for template data structure
                     rubric_template_data[domain_name] = {"title": domain_title, "elements": []}

                     # Display Domain Header
                     st.markdown(f"<div style='background-color:{background};padding:12px;border-radius:10px;margin-bottom:5px;'><h4 style='margin-bottom:5px;'>{domain_name}: {domain_title}</h4></div>", unsafe_allow_html=True)

                     col_label_template = 'B' # Column where element labels are in template
                     col_descriptors_start = 'C' # Column where descriptor for rating 6 starts
                     row_start_template = int(start_cell[1:])

                     for i in range(count):
                         element_number = f"{idx+1}.{i+1}"
                         element_row_template = row_start_template + i

                         # Read Element Label from template
                         try:
                             label = rubric_template_ws[f"{col_label_template}{element_row_template}"].value or f"Element {element_number}"
                         except Exception:
                             label = f"Element {element_number}" # Fallback label

                         # Read Descriptors from template (assuming C-H for 6 down to 1)
                         descriptors = {}
                         descriptor_text_full = ""
                         for j in range(6): # Assuming ratings 6 down to 1
                             descriptor_col = get_column_letter(ord(col_descriptors_start) - ord('A') + j + 1) # C, D, E, F, G, H
                             rating_level = 6 - j
                             try:
                                 desc_value = rubric_template_ws[f"{descriptor_col}{element_row_template}"].value
                                 if desc_value is not None:
                                     descriptors[rating_level] = str(desc_value)
                                     # Also build formatted markdown for expander
                                     descriptor_text_full += f"**{rating_level}:** {desc_value}\n\n"
                             except Exception:
                                 pass # Ignore if reading descriptor fails

                         # Store element info from template for PDF/Feedback structure
                         rubric_template_data[domain_name]["elements"].append({
                             "number": element_number,
                             "label": label,
                             "descriptors": descriptors, # Store all descriptors
                             "descriptor_text_full": descriptor_text_full.strip() # Store formatted text for expander
                         })

                         st.markdown(f"<div style='background-color:{background};padding:8px;border-radius:6px;'>", unsafe_allow_html=True)
                         st.markdown(f"**{element_number} – {label}**")

                         # Rubric Guidance Expander
                         with st.expander(strings["expander_rubric_descriptors"]):
                             if descriptor_text_full:
                                 st.markdown(descriptor_text_full)
                             else:
                                 st.info(strings["info_no_descriptors"])

                         # Input Widgets - Rating and Notes
                         col1, col2 = st.columns([1, 2])
                         with col1:
                             # Use session state for the value and unique key
                             rating_key = f"{domain_name}_{i}_rating" # Key for the rating selectbox
                             current_rating = st.session_state.get(rating_key, "NA") # Default to "NA" for new, load for existing
                             # Ensure the default index logic handles the current_rating value correctly
                             rating_options = [6, 5, 4, 3, 2, 1, "NA"]
                             try:
                                 initial_rating_index = rating_options.index(current_rating)
                             except ValueError:
                                 initial_rating_index = rating_options.index("NA") # Default to index of "NA" if current value is not in options

                             st.selectbox(
                                 strings["label_rating_for"].format(element_number),
                                 options=rating_options,
                                 index=initial_rating_index,
                                 key=rating_key # Use the key from session state
                             )
                              # Value is automatically updated in st.session_state[rating_key]

                         with col2:
                             # Use session state for the value and unique key
                             note_key = f"{domain_name}_{i}_note" # Key for the notes text area
                             current_note = st.session_state.get(note_key, "") # Default to empty string for new, load for existing
                             st.text_area(
                                 strings["label_write_notes"].format(element_number), # Use localized string
                                 value=current_note,
                                 key=note_key, # Use the key from session state
                                 height=100
                             )
                              # Value is automatically updated in st.session_state[note_key]

                         st.markdown("</div>", unsafe_allow_html=True) # Close the div for the element background

             except KeyError:
                 st.error(strings["error_template_not_found"]) # "LO 1" sheet not found
                 # Prevent further execution if template is missing
                 st.stop()
             except Exception as e:
                 st.error(f"Error reading rubric details from template: {e}")
                 # Prevent further execution if template reading fails
                 st.stop()

             # Overall Notes (Integrated from snippet 2)
             overall_notes = st.text_area(strings.get("label_overall_notes", "General Notes for this Lesson Observation"), value=st.session_state.get('overall_notes', ''), key='overall_notes') # Added string lookup, session state

             # Feedback Checkbox (Placed before the button so state is captured on click)
             send_feedback = st.checkbox(strings["checkbox_send_feedback"], value=st.session_state.get('checkbox_send_feedback', False), key='checkbox_send_feedback')

             # --- Save Button ---
             # The save button starts the process of writing data, calculating scores, and generating PDF/Log
             if st.button(strings["button_save_observation"]):

                 # --- Validation ---
                 # Check essential fields from session state
                 essential_filled = True
                 required_keys = ['observer_name', 'teacher_name', 'school_name', 'grade', 'subject', 'gender', 'observation_date']
                 for key in required_keys:
                     if not st.session_state.get(key):
                         essential_filled = False
                         break # Exit loop early if any essential field is empty

                 if not essential_filled:
                     st.warning(strings["warning_fill_basic_info"])
                     st.stop() # Stop execution if validation fails

                 # Validate numeric fields if necessary (e.g., students, males, females) from session state
                 num_students, num_males, num_females = None, None, None
                 try:
                     # Attempt conversion to int, handle empty strings as None
                     students_val = st.session_state.get('students', '')
                     males_val = st.session_state.get('males', '')
                     females_val = st.session_state.get('females', '')

                     num_students = int(students_val) if students_val else None
                     num_males = int(males_val) if males_val else None
                     num_females = int(females_val) if females_val else None

                     # Add more specific validation if needed (e.g., males + females == students if all are provided)

                 except ValueError:
                     st.warning(strings["warning_numeric_fields"]) # Use localized string
                     st.stop()

                 # --- Get Target Sheet for Saving ---
                 target_sheet_name = st.session_state.get('target_sheet_name')
                 ws_to_save = None # Initialize worksheet for saving

                 # Determine if we need to create a new sheet now
                 if target_sheet_name == strings["option_create_new"] or (target_sheet_name and target_sheet_name not in wb.sheetnames):
                     # Logic to create the new sheet
                     if "LO 1" in wb.sheetnames:
                         try:
                             # Determine the actual new sheet name (LO X) - Find highest existing number + 1
                             existing_lo_numbers = []
                             for sheet in wb.sheetnames:
                                 if sheet.startswith("LO "):
                                     suffix = sheet[3:].strip() # Get suffix and strip whitespace
                                     if suffix.isdigit():
                                         existing_lo_numbers.append(int(suffix))

                             next_index = max(existing_lo_numbers) + 1 if existing_lo_numbers else 1
                             sheet_name_to_save_actual = f"LO {next_index}"

                             # Copy template and set title
                             ws_to_save = wb.copy_worksheet(wb["LO 1"])
                             ws_to_save.title = sheet_name_to_save_actual
                             st.success(strings["success_sheet_created"].format(sheet_name_to_save_actual))
                              # Update session state with the real sheet name for subsequent saves/reruns
                             st.session_state['target_sheet_name'] = sheet_name_to_save_actual
                             st.session_state['current_sheet_name'] = sheet_name_to_save_actual # Update selector state

                         except Exception as e:
                             st.error(f"Error creating new sheet for saving: {e}")
                             st.stop() # Cannot save if sheet creation failed
                     else:
                         st.error(strings["error_template_not_found"]) # Template missing, cannot create
                         st.stop()
                 elif target_sheet_name:
                     # Use the existing selected sheet for saving
                     sheet_name_to_save_actual = target_sheet_name
                     try:
                         ws_to_save = wb[sheet_name_to_save_actual]
                     except KeyError:
                         st.error(f"Error: Selected sheet '{sheet_name_to_save_actual}' not found in workbook.")
                         st.stop()
                     except Exception as e:
                         st.error(f"Error accessing selected sheet '{sheet_name_to_save_actual}': {e}")
                         st.stop()
                 else:
                     st.error("No target sheet determined for saving.")
                     st.stop() # Should not happen if target_sheet_name is set on sheet selection

                 # Proceed with writing and saving only if we have a worksheet to save to
                 if ws_to_save:
                     # --- Write Data to Excel Sheet (ws_to_save) ---
                     try:
                         # Basic Info from session state
                         ws_to_save["B5"].value = st.session_state.get('gender')
                         ws_to_save["B6"].value = num_students # Use validated numbers
                         ws_to_save["B7"].value = num_males
                         ws_to_save["B8"].value = num_females
                         ws_to_save["D2"].value = st.session_state.get('subject')
                         # Recalculate duration display string for saving
                         minutes_save = 0
                         duration_label_save = "N/A"
                         time_in_ss = st.session_state.get('time_in')
                         time_out_ss = st.session_state.get('time_out')
                         if time_in_ss is not None and time_out_ss is not None:
                             dummy_date = date.today()
                             start_dt = datetime.combine(dummy_date, time_in_ss)
                             end_dt = datetime.combine(dummy_date, time_out_ss)
                             if end_dt < start_dt: end_dt += timedelta(days=1)
                             lesson_duration_save = end_dt - start_dt
                             minutes_save = round(lesson_duration_save.total_seconds() / 60)
                             duration_label_save = strings["duration_full_lesson"] if minutes_save >= 40 else strings["duration_walkthrough"]

                         ws_to_save["D3"].value = duration_label_save # Save calculated duration label
                         ws_to_save["D4"].value = st.session_state.get('period')
                         # Save time objects directly, Excel should format them correctly
                         ws_to_save["D7"].value = st.session_state.get('time_in')
                         ws_to_save["D8"].value = st.session_state.get('time_out')

                         # Observation Date - Assuming D10, adjust cell if needed
                         ws_to_save["D10"].value = st.session_state.get('observation_date') # Save date object

                         # Metadata from session state
                         ws_to_save["Z1"].value = "Observer Name" # Keep header, redundant but matches template
                         ws_to_save["AA1"].value = st.session_state.get('observer_name')
                         ws_to_save["Z2"].value = "Teacher Observed"
                         ws_to_save["AA2"].value = st.session_state.get('teacher_name')
                         ws_to_save["Z3"].value = "Observation Type"
                         ws_to_save["AA3"].value = st.session_state.get('observation_type')
                         ws_to_save["Z4"].value = "Timestamp"
                         ws_to_save["AA4"].value = datetime.now() # Save current datetime object
                         ws_to_save["Z5"].value = "Operator"
                         ws_to_save["AA5"].value = st.session_state.get('operator')
                         ws_to_save["Z6"].value = "School Name"
                         ws_to_save["AA6"].value = st.session_state.get('school_name')
                         ws_to_save["Z7"].value = "General Notes"
                         ws_to_save["AA7"].value = st.session_state.get('overall_notes')
                         ws_to_save["Z8"].value = "Teacher Email" # Added header for email
                         ws_to_save["AA8"].value = st.session_state.get('teacher_email') # Save teacher email


                         # Rubric Scores and Notes - Write values from session state to sheet
                         rubric_domains_structure = { # Re-define or access if defined globally
                              "Domain 1": ("I11", 5), "Domain 2": ("I20", 3), "Domain 3": ("I27", 4), "Domain 4": ("I35", 3),
                              "Domain 5": ("I42", 2), "Domain 6": ("I48", 2), "Domain 7": ("I54", 2), "Domain 8": ("I60", 3), "Domain 9": ("I67", 2)
                         }

                         # Prepare data structure for PDF, combining template info with user inputs
                         pdf_rubric_data = {"domain_data": {}}

                         # Dictionaries to store scores, notes, and descriptors for feedback generation and PDF
                         domain_calculated_averages = {}
                         overall_score = None

                         for idx, (domain_name, (start_cell, count)) in enumerate(rubric_domains_structure.items()):
                              col_rating_save = start_cell[0] # Column where ratings are saved (e.g., 'I')
                              col_note_save = 'J' # Column for notes (based on snippet 2)
                              row_start_save = int(start_cell[1:])

                              domain_elements_for_pdf = []
                              numeric_element_scores_in_domain = [] # For Python average calculation

                              # Find the template info for this domain
                              # Use .get() with a default to handle cases where domain_name might not be in template_data (shouldn't happen with correct structure)
                              domain_template_info = rubric_template_data.get(domain_name, {"title": domain_name, "elements": []})
                              pdf_rubric_data["domain_data"][domain_name] = {
                                   "title": domain_template_info.get("title", domain_name), # Use .get() for safety
                                   "elements": [] # Will populate with elements + ratings/notes
                              }

                              for i in range(count):
                                   row_save = row_start_save + i
                                   rating_key = f"{domain_name}_{i}_rating"
                                   note_key = f"{domain_name}_{i}_note"

                                   # Get values from session state (user inputs)
                                   rating_value = st.session_state.get(rating_key, "NA")
                                   note_value = st.session_state.get(note_key, "")

                                   # Write values to the sheet
                                   ws_to_save[f"{col_rating_save}{row_save}"].value = rating_value
                                   ws_to_save[f"{col_note_save}{row_save}"].value = note_value

                                   # Collect numeric scores for Python calculation
                                   if isinstance(rating_value, (int, float)):
                                        numeric_element_scores_in_domain.append(float(rating_value))


                                   # Find the element details from the template structure for the PDF
                                   element_details_template = next((item for item in domain_template_info.get("elements", []) if item["number"] == f"{idx+1}.{i+1}"), None)

                                   if element_details_template:
                                        # Get the specific descriptor text for the chosen rating
                                        descriptor_for_rating = element_details_template["descriptors"].get(rating_value, "") if isinstance(rating_value, int) else ""

                                        # Add element info (from template + user inputs) to PDF data structure
                                        pdf_rubric_data["domain_data"][domain_name]["elements"].append({
                                             'label': element_details_template.get("label", "Unknown Element"), # Use .get() for safety
                                             'rating': rating_value,
                                             'note': note_value,
                                             'descriptor': descriptor_for_rating, # Pass the specific descriptor text
                                             'number': element_details_template.get("number", f"{idx+1}.{i+1}") # Use .get() for safety
                                        })


                              # --- Write Excel formulas for domain averages and judgments ---
                              # These formulas work directly in Excel. We also calculate in Python for feedback/PDF.
                              score_range = f"{col_rating_save}{row_start_save}:{col_rating_save}{row_start_save + count - 1}"
                              avg_cell = f"{col_rating_save}{row_start_save + count}"
                              judgment_cell = f"K{row_start_save + count}" # Adjusted judgment column to K, adjust if needed

                              # Write formulas to the sheet
                              # Use AGGREGATE to ignore errors (#DIV/0!) in the range
                              ws_to_save[avg_cell].value = f'=IFERROR(AGGREGATE(1, 6, {score_range}), "")' # AGGREGATE(1, 6, range) is AVERAGE ignoring errors/hidden rows
                              ws_to_save[judgment_cell].value = f'=IF({avg_cell}="","",IF({avg_cell}>=5.5,"{strings["perf_level_excellent"]}",IF({avg_cell}>=4.5,"{strings["perf_level_good"]}",IF({avg_cell}>=3.5,"{strings["perf_level_acceptable"]}",IF({avg_cell}>=2.5,"{strings["perf_level_weak"]}","{strings["perf_level_very_weak"]}") ))))'


                              # --- Calculate Python Averages and Judgments for Feedback/PDF ---
                              if numeric_element_scores_in_domain:
                                   domain_avg = statistics.mean(numeric_element_scores_in_domain)
                                   domain_calculated_averages[domain_name] = domain_avg
                                   # Add calculated average and judgment to PDF data structure
                                   pdf_rubric_data["domain_data"][domain_name]["average"] = domain_avg
                                   pdf_rubric_data["domain_data"][domain_name]["judgment"] = get_performance_level(domain_avg, strings)
                              else:
                                   domain_calculated_averages[domain_name] = None
                                   pdf_rubric_data["domain_data"][domain_name]["average"] = None
                                   pdf_rubric_data["domain_data"][domain_name]["judgment"] = strings["overall_score_na"]


                         # --- Calculate Overall Python Average and Judgment ---
                         all_numeric_domain_scores = []
                         for d_name, d_avg in domain_calculated_averages.items():
                              if d_avg is not None:
                                   all_numeric_domain_scores.append(d_avg) # Use domain averages for overall average

                         if all_numeric_domain_scores:
                              overall_score = statistics.mean(all_numeric_domain_scores)
                              overall_judgment = get_performance_level(overall_score, strings)
                              overall_score_display = strings["overall_score_value"].format(overall_score)
                         else:
                              overall_score = None
                              overall_judgment = strings["overall_score_na"]
                              overall_score_display = strings["overall_score_na"]


                         # Add overall score, judgment, and notes to the PDF data structure
                         pdf_rubric_data["overall_score"] = overall_score
                         pdf_rubric_data["overall_score_display"] = overall_score_display # Add formatted score for PDF table
                         pdf_rubric_data["overall_judgment"] = overall_judgment
                         pdf_rubric_data["overall_notes"] = st.session_state.get('overall_notes', '') # Get overall notes from session state
                         pdf_rubric_data["observer_name"] = st.session_state.get('observer_name', '') # Add basic info to PDF data
                         pdf_rubric_data["teacher_name"] = st.session_state.get('teacher_name', '')
                         pdf_rubric_data["operator"] = st.session_state.get('operator', '')
                         pdf_rubric_data["school_name"] = st.session_state.get('school_name', '')
                         pdf_rubric_data["grade"] = st.session_state.get('grade', '')
                         pdf_rubric_data["subject"] = st.session_state.get('subject', '')
                         pdf_rubric_data["gender"] = st.session_state.get('gender', '')
                         pdf_rubric_data["students"] = st.session_state.get('students', '')
                         pdf_rubric_data["males"] = st.session_state.get('males', '')
                         pdf_rubric_data["females"] = st.session_state.get('females', '')
                         pdf_rubric_data["observation_date"] = st.session_state.get('observation_date', '')
                         pdf_rubric_data["time_in"] = st.session_state.get('time_in', '')
                         pdf_rubric_data["time_out"] = st.session_state.get('time_out', '')
                         pdf_rubric_data["period"] = st.session_state.get('period', '')
                         pdf_rubric_data["observation_type"] = st.session_state.get('observation_type', '')
                         pdf_rubric_data["duration_display"] = f"{minutes_save} minutes ({duration_label_save})" # Add formatted duration for PDF table


                         # --- Update Observation Log ---
                         log_sheet_name = strings["feedback_log_sheet_name"]
                         if log_sheet_name not in wb.sheetnames:
                              log_ws = wb.create_sheet(log_sheet_name)
                              log_ws.append(strings["feedback_log_header"]) # Use headers from strings dictionary
                         else:
                              try:
                                   log_ws: Worksheet = wb[log_sheet_name]
                                   # Ensure headers exist if sheet was empty or different
                                   if log_ws.max_row == 0 or log_ws['A1'].value not in strings["feedback_log_header"]: # Check if empty or wrong header
                                        log_ws.append(strings["feedback_log_header"])
                              except Exception as e:
                                   st.warning(f"Could not access or validate log sheet '{log_sheet_name}', attempting to create new. Error: {e}")
                                   log_ws = wb.create_sheet(log_sheet_name + "_new") # Create with a different name to avoid conflict
                                   log_sheet_name = log_sheet_name + "_new" # Update name for messages


                         # Check if an entry for this sheet already exists in the log
                         existing_row_index = None
                         for row_index in range(2, log_ws.max_row + 1): # Start from row 2 to skip header
                              if log_ws[f'A{row_index}'].value == sheet_name_to_save_actual:
                                   existing_row_index = row_index
                                   break

                         # Prepare data for the log row based on the headers
                         log_row_data = {
                              "Sheet": sheet_name_to_save_actual, # Use the actual sheet name saved
                              "Observer": st.session_state.get('observer_name', ''),
                              "Teacher": st.session_state.get('teacher_name', ''),
                              "Email": st.session_state.get('teacher_email', ''),
                              "School": st.session_state.get('school_name', ''),
                              "Subject": st.session_state.get('subject', ''),
                              "Date": st.session_state.get('observation_date').strftime("%Y-%m-%d") if st.session_state.get('observation_date') else "",
                              "Overall Judgment": overall_judgment, # Include overall judgment
                              "Overall Score": overall_score if overall_score is not None else strings["overall_score_na"], # Include overall score
                              "Summary Notes": st.session_state.get('overall_notes', '') # Include overall notes
                         }
                         # Ensure the order matches feedback_log_header
                         ordered_log_row = [log_row_data.get(header, "") for header in strings["feedback_log_header"]]

                         # Append or Update the row
                         try:
                              if existing_row_index:
                                   # Update existing row
                                   for col_index, value in enumerate(ordered_log_row):
                                        log_ws.cell(row=existing_row_index, column=col_index + 1).value = value
                              else:
                                   # Append new row
                                   log_ws.append(ordered_log_row)

                              st.success(strings["success_feedback_log_updated"]) # Use localized string
                         except Exception as e:
                              st.error(strings["error_updates
