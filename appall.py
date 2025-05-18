#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon May 12 00:42:10 2025

@author: paulaassaf
"""

import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime, timedelta, date, time # Import time specifically
import os
import statistics
import pandas as pd
import matplotlib.pyplot as plt
import csv
import io
from openpyxl.utils import get_column_letter

# Import ReportLab modules for PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
import re # Import regex for cleaning HTML tags
import numpy as np # Import numpy for isnan

# --- Set Streamlit Page Config (MUST BE THE FIRST STREAMLIT COMMAND) ---
st.set_page_config(page_title="Lesson Observation Tool", layout="wide")

# --- Logo File Paths ---
# Define a dictionary mapping school names to logo file paths
# Ensure these paths are correct relative to your script location
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
    # Add other school logos here
}

# --- Text Strings for Localization ---
en_strings = {
    "page_title": "Lesson Observation Tool",
    "sidebar_select_page": "Choose a page:",
    "page_lesson_input": "Lesson Observation Input",
    "page_analytics": "Lesson Observation Analytics",
    "page_help": "App Information and Guidelines",
    "title_lesson_input": "Weekly Lesson Observation Input Tool",
    "title_help": "App Information and Guidelines",
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
    "warning_could_not_calculate_duration": "Could not calculate lesson duration.",
    "label_period": "Period",
    "label_obs_type": "Observation Type",
    "option_individual": "Individual",
    "option_joint": "Joint",
    "subheader_rubric_scores": "Rubric Scores",
    "expander_rubric_descriptors": "Rubric Guidance",
    "info_no_descriptors": "No rubric guidance available.",
    "label_rating_for": "Rating for {}",
    "label_write_notes": "Write notes for {}",
    "checkbox_send_feedback": "✉️ Generate Feedback Report (for PDF)",
    "button_save_observation": "💾 Save Observation",
    "warning_fill_essential": "Please fill in all essential information before saving.",
    "success_data_saved": "Observation data saved to workbook.",
    "error_saving_workbook": "Error saving workbook:",
    "download_workbook": "📥 Download updated workbook",
    "feedback_subject": "Lesson Observation Feedback",
    "feedback_greeting": "Dear {},\n\nYour lesson observation from {} has been saved.\n\n",
    "feedback_observer": "Observer: {}\n",
    "feedback_duration": "Duration: {}\n",
    "feedback_subject_fb": "Subject: {}\n",
    "feedback_school": "School: {}\n\n",
    "feedback_summary_header": "Here is a summary of your ratings based on the rubric:\n\n",
    "feedback_domain_header": "**{}: {}**\n",
    "feedback_element_rating": "- **{}:** Rating **{}**\n",
    "feedback_descriptor_for_rating": "  *Guidance for rating {}:* {}\n",
    "feedback_overall_score": "\n**Overall Average Score:** {:.2f}\n\n",
    "feedback_domain_average": "  *Domain Average:* {:.2f}\n",
    "feedback_performance_summary": "**Performance Summary:**\n",
    "overall_performance_level_text": "Overall Performance Level: {}",
    "feedback_domain_performance": "{}: {}\n",
    "feedback_support_plan_intro": "\n**Support Plan Recommended:**\n",
    "feedback_next_steps_intro": "\n**Suggested Next Steps:**\n",
    "feedback_closing": "\nBased on these ratings, please review your updated workbook for detailed feedback and areas for development.\n\n",
    "feedback_regards": "Regards,\nSchool Leadership Team",
    "success_feedback_generated": "Feedback generated.",
    "success_feedback_log_updated": "Feedback log updated.",
    "error_updating_log": "Error updating feedback log in workbook:",
    "title_analytics": "Lesson Observation Analytics Dashboard",
    "warning_no_lo_sheets_analytics": "No 'LO ' sheets found in the workbook for analytics.",
    "subheader_avg_score_overall": "Average Score per Domain (Across all observations)",
    "info_no_numeric_scores_overall": "No numeric scores found across all observations to calculate overall domain averages.",
    "subheader_data_summary": "Observation Data Summary",
    "subheader_filter_analyze": "Filter and Analyze",
    "filter_by_school": "Filter by School",
    "filter_by_grade": "Filter by Grade",
    "filter_by_subject": "Filter by Subject",
    "filter_by_operator": "Filter by Operator",
    "filter_by_observer_an": "Filter by Observer",
    "option_all": "All",
    "subheader_avg_score_filtered": "Average Score per Domain (Filtered)",
    "info_no_numeric_scores_filtered": "No observations matching the selected filters contain numeric scores for domain averages.",
    "subheader_observer_distribution": "Observer Distribution (Filtered)",
    "info_no_observer_data_filtered": "No observer data found for the selected filters.",
    "info_no_observation_data_filtered": "No observation data found for the selected filters.",
    "error_loading_analytics": "Error loading or processing workbook for analytics:",
    "overall_score_label": "Overall Score:",
    "overall_score_value": "**{:.2f}**",
    "overall_score_na": "**N/A**",
    "arabic_toggle_label": "عرض باللغة العربية (Display in Arabic)",
    "feedback_log_sheet_name": "Feedback Log", # Using English name for robustness with code logic
    "feedback_log_header": ["Sheet", "Observer", "Teacher", "Email", "School", "Subject", "Date", "Overall Judgment", "Overall Score", "Summary Notes"],
    "download_feedback_log_csv": "📥 Download Feedback Log (CSV)",
    "error_generating_log_csv": "Error generating Feedback Log CSV:",
    "download_overall_avg_csv": "📥 Download Overall Domain Averages (CSV)",
    "download_overall_avg_excel": "📥 Download Overall Domain Averages (Excel)",
    "download_filtered_avg_csv": "📥 Download Filtered Domain Averages (CSV)", # Not currently generated as a separate file
    "download_filtered_avg_excel": "📥 Download Filtered Domain Averages (Excel)", # Not currently generated as a separate file
    "download_filtered_data_csv": "📥 Download Filtered Visit Data (CSV)",
    "download_filtered_data_excel": "📥 Download Filtered Visit Data (Excel)",
    "label_observation_date": "Observation Date",
    "filter_start_date": "Start Date",
    "filter_end_date": "End Date",
    "filter_teacher": "Filter by Teacher",
    "subheader_teacher_performance": "Teacher Performance Over Time",
    "info_select_teacher": "Select a teacher to view individual performance analytics.",
    "info_no_obs_for_teacher": "No observations found for the selected teacher within the applied filters.",
    "subheader_teacher_domain_trend": "{} Domain Performance Trend",
    "subheader_teacher_overall_avg": "{} Overall Average Score (Filtered)",
    "perf_level_very_weak": "Very Weak",
    "perf_level_weak": "Weak",
    "perf_level_acceptable": "Acceptable",
    "perf_level_good": "Good",
    "perf_level_excellent": "Excellent",
    "plan_very_weak_overall": "Overall performance is Very Weak. A comprehensive support plan is required. Focus on fundamental teaching practices such as classroom management, lesson planning, and basic instructional strategies. Seek guidance from your mentor teacher and school leadership.",
    "plan_weak_overall": "Overall performance is Weak. A support plan is recommended. Identify 1-2 key areas for improvement from the observation and work with your mentor teacher to develop targeted strategies. Consider observing experienced colleagues in these areas.",
    "plan_weak_domain": "Performance in **{}** is Weak. Focus on developing skills related to: {}. Suggested actions include: [Specific action 1], [Specific action 2].", # Placeholder actions
    "steps_acceptable_overall": "Overall performance is Acceptable. Continue to build on your strengths. Identify one area for growth to refine your practice and enhance student learning.",
    "steps_good_overall": "Overall performance is Good. You demonstrate effective teaching practices. Explore opportunities to share your expertise with colleagues, perhaps through informal mentoring or presenting successful strategies.",
    "steps_good_domain": "Performance in **{}** is Good. You demonstrate strong skills in this area. Consider exploring advanced strategies related to: {}. Suggested actions include: [Specific advanced action 1], [Specific advanced action 2].", # Placeholder actions
    "steps_excellent_overall": "Overall performance is Excellent. You are a role model for effective teaching. Consider leading professional development sessions or mentoring less experienced teachers.",
    "steps_excellent_domain": "Performance in **{}** is Excellent. Your practice in this area is exemplary. Continue to innovate and refine your practice, perhaps by researching and implementing cutting-edge strategies related to: {}.", # Placeholder actions
    "no_specific_plan_needed": "Performance is at an acceptable level or above. No immediate support plan required based on this observation. Focus on continuous improvement based on your professional goals.",
    "warning_fill_essential": "Please fill in all essential information before saving.", # Simplified generic warning
    "warning_numeric_fields": "Please enter valid numbers for Students, Males, and Females.",
    "success_pdf_generated": "Feedback PDF generated successfully.",
    "download_feedback_pdf": "📥 Download Feedback PDF",
    "checkbox_cleanup_sheets": "🪟 Clean up unused LO sheets (no observer name)",
    "warning_sheets_removed": "Removed {} unused LO sheets.",
    "info_reloaded_workbook": "Reloaded workbook after cleanup.",
    "info_no_sheets_to_cleanup": "No unused LO sheets found to clean up.",
    "expander_guidelines": "📘 Click here to view observation guidelines",
    "info_no_guidelines": "Guidelines sheet is empty or could not be read.",
    "warning_select_create_sheet": "Please select or create a valid sheet to proceed.",
    "label_overall_notes": "General Notes for this Lesson Observation",
}

# Placeholder Arabic strings - REPLACE THESE WITH ACTUAL TRANSLATIONS
# NOTE: Arabic support in ReportLab PDFs requires Arabic fonts and potentially bi-directional text handling.
# The translations below are placeholders and ReportLab PDF output for these may not render correctly without additional setup.
ar_strings = {
    "page_title": "أداة تقييم الزيارات الصفية",
    "sidebar_select_page": "اختر صفحة:",
    "page_lesson_input": "إدخال تقييم الزيارة",
    "page_analytics": "تحليلات الزيارات الصفية",
    "page_help": "معلومات وإرشادات التطبيق",
    "title_lesson_input": "أداة إدخال زيارة صفية أسبوعية",
    "title_help": "معلومات وإرشادات التطبيق",
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
    "expander_rubric_descriptors": "إرشادات الدليل",
    "info_no_descriptors": "لا توجد إرشادات دليل متاحة لهذا العنصر.",
    "label_rating_for": "التقييم لـ {}",
    "label_write_notes": "كتابة ملاحظات لـ {}",
    "checkbox_send_feedback": "✉️ إنشاء تقرير الملاحظات (للملف PDF)",
    "button_save_observation": "💾 حفظ الزيارة",
    "warning_fill_essential": "يرجى ملء جميع حقول المعلومات الأساسية قبل الحفظ.",
    "success_data_saved": "تم حفظ بيانات الزيارة في المصنف.",
    "error_saving_workbook": "خطأ في حفظ المصنف:",
    "download_workbook": "📥 تنزيل المصنف المحدث",
    "feedback_subject": "ملاحظات الزيارة الصفية",
    "feedback_greeting": "عزيزي {},\n\nتم حفظ زيارتك الصفية من {}.\n\n",
    "feedback_observer": "المراقب: {}\n",
    "feedback_duration": "المدة: {}\n",
    "feedback_subject_fb": "المادة: {}\n",
    "feedback_school": "المدرسة: {}\n\n",
    "feedback_summary_header": "إليك ملخص لتقييماتك بناءً على الدليل:\n\n",
    "feedback_domain_header": "**{}: {}**\n",
    "feedback_element_rating": "- **{}:** التقييم **{}**\n",
    "feedback_descriptor_for_rating": "  *إرشادات للتقييم {}:* {}\n",
    "feedback_overall_score": "\n**متوسط الدرجة الإجمالي:** {:.2f}\n\n",
    "feedback_domain_average": "  *متوسط المجال:* {:.2f}\n",
    "feedback_performance_summary": "**ملخص الأداء:**\n",
    "overall_performance_level_text": "مستوى الأداء الإجمالي: {}",
    "feedback_domain_performance": "{}: {}\n",
    "feedback_support_plan_intro": "\n**خطة الدعم الموصى بها:**\n",
    "feedback_next_steps_intro": "\n**الخطوات التالية المقترحة:**\n",
    "feedback_closing": "\nبناءً على هذه التقييمات، يرجى مراجعة المصنف المحدث للحصول على ملاحظات تفصيلية ومجالات التطوير.\n\n",
    "feedback_regards": "مع التحيات,\nفريق قيادة المدرسة",
    "success_feedback_generated": "تم إنشاء الملاحظات.",
    "success_feedback_log_updated": "تم تحديث سجل الملاحظات.",
    "error_updating_log": "خطأ في تحديث سجل الملاحظات في المصنف:",
    "title_analytics": "لوحة تحليلات الزيارة الصفية",
    "warning_no_lo_sheets_analytics": "لم يتم العثور على أوراق 'LO ' في المصنف للتحليلات.",
    "subheader_avg_score_overall": "متوسط الدرجة لكل مجال (عبر جميع الزيارات)",
    "info_no_numeric_scores_overall": "لم يتم العثور على درجات رقمية عبر جميع الزيارات لحساب متوسطات المجال الإجمالية.",
    "subheader_data_summary": "ملخص بيانات الزيارة",
    "subheader_filter_analyze": "تصفية وتحليل",
    "filter_by_school": "تصفية حسب المدرسة",
    "filter_by_grade": "تصفية حسب الصف",
    "filter_by_subject": "تصفية حسب المادة",
    "filter_by_operator": "تصفية حسب المشغل",
    "filter_by_observer_an": "تصفية حسب المراقب",
    "option_all": "الكل",
    "subheader_avg_score_filtered": "متوسط الدرجة لكل مجال (مصفى)",
    "info_no_numeric_scores_filtered": "لا توجد زيارات مطابقة للمرشحات المحددة تحتوي على درجات رقمية لمتوسطات المجال.",
    "subheader_observer_distribution": "توزيع المراقبين (مصفى)",
    "info_no_observer_data_filtered": "لم يتم العثور على بيانات المراقب للمرشحات المحددة.",
    "info_no_observation_data_filtered": "لم يتم العثور على بيانات الزيارة للمرشحات المحددة.",
    "error_loading_analytics": "خطأ في تحميل أو معالجة المصنف للتحليلات:",
    "overall_score_label": "النتيجة الإجمالية:",
    "overall_score_value": "**{:.2f}**",
    "overall_score_na": "**غير متوفر**",
    "arabic_toggle_label": "عرض باللغة العربية (Display in Arabic)",
    "feedback_log_sheet_name": "سجل الملاحظات", # Using Arabic name for robustness
    "feedback_log_header": ["Sheet", "Observer", "Teacher", "Email", "School", "Subject", "Date", "Overall Judgment", "Overall Score", "Summary Notes"], # Keep English keys for code, display Arabic in UI if needed
    "download_feedback_log_csv": "📥 تنزيل سجل الملاحظات (CSV)",
    "error_generating_log_csv": "خطأ في إنشاء سجل الملاحظات CSV:",
    "download_overall_avg_csv": "📥 تنزيل متوسطات المجال الإجمالية (CSV)",
    "download_overall_avg_excel": "📥 تنزيل متوسطات المجال الإجمالية (Excel)",
    "download_filtered_avg_csv": "📥 تنزيل متوسطات المجال المصفاة (CSV)", # Not currently generated as a separate file
    "download_filtered_avg_excel": "📥 تنزيل متوسطات المجال المصفاة (Excel)", # Not currently generated as a separate file
    "download_filtered_data_csv": "📥 تنزيل بيانات الزيارة المصفاة (CSV)",
    "download_filtered_data_excel": "📥 تنزيل بيانات الزيارة المصفاة (Excel)",
    "label_observation_date": "تاريخ الزيارة",
    "filter_start_date": "تاريخ البدء",
    "filter_end_date": "تاريخ الانتهاء",
    "filter_teacher": "تصفية حسب المعلم",
    "subheader_teacher_performance": "أداء المعلم بمرور الوقت",
    "info_select_teacher": "حدد معلمًا لعرض تحليلات الأداء الفردي.",
    "info_no_obs_for_teacher": "لم يتم العثور على زيارات للمعلم المحدد ضمن المرشحات المطبقة.",
    "subheader_teacher_domain_trend": "اتجاه أداء مجال {}",
    "subheader_teacher_overall_avg": "متوسط الدرجة الإجمالي لـ {} (مصفى)",
    "perf_level_very_weak": "ضعيف جداً",
    "perf_level_weak": "ضعيف",
    "perf_level_acceptable": "مقبول",
    "perf_level_good": "جيد",
    "perf_level_excellent": "ممتاز",
    "plan_very_weak_overall": "الأداء الإجمالي ضعيف جداً. تتطلب خطة دعم شاملة. ركز على الممارسات التعليمية الأساسية مثل إدارة الصف، وتخطيط الدرس، والاستراتيجيات التعليمية الأساسية. اطلب التوجيه من معلمك الموجه وقيادة المدرسة.",
    "plan_weak_overall": "الأداء الإجمالي ضعيف. يوصى بخطة دعم. حدد 1-2 من المجالات الرئيسية للتحسين من الملاحظة واعمل مع معلمك الموجه لتطوير استراتيجيات مستهدفة. فكر في ملاحظة الزملاء ذوي الخبرة في هذه المجالات.",
    "plan_weak_domain": "الأداء في **{}** ضعيف. ركز على تطوير المهارات المتعلقة بـ: {}. الإجراءات المقترحة تشمل: [إجراء محدد 1]، [إجراء محدد 2].",
    "steps_acceptable_overall": "الأداء الإجمالي مقبول. استمر في البناء على نقاط قوتك. حدد مجالًا واحدًا للنمو لتحسين ممارستك وتعزيز تعلم الطلاب.",
    "steps_good_overall": "الأداء الإجمالي جيد. أنت تظهر ممارسات تعليمية فعالة. استكشف فرص مشاركة خبرتك مع الزملاء، ربما من خلال التوجيه غير الرسمي أو تقديم استراتيجيات ناجحة.",
    "steps_good_domain": "الأداء في **{}** جيد. أنت تظهر مهارات قوية في هذا المجال. فكر في استكشاف استراتيجيات متقدمة تتعلق بـ: {}. الإجراءات المقترحة تشمل: [إجراء متقدم محدد 1]، [إجراء متقدم محدد 2].",
    "steps_excellent_overall": "الأداء الإجمالي ممتاز. أنت نموذج يحتذى به في التدريس الفعال. فكر في قيادة جلسات التطوير المهني أو توجيه المعلمين الأقل خبرة.",
    "steps_excellent_domain": "الأداء في **{}** ممتاز. ممارستك في هذا المجال نموذجية. استمر في الابتكار وتحسين ممارستك، ربما من خلال البحث وتطبيق استراتيجيات حديثة تتعلق بـ: {}.",
    "no_specific_plan_needed": "الأداء عند مستوى مقبول أو أعلى. لا توجد خطة دعم فورية مطلوبة بناءً على هذه الملاحظة. ركز على التحسين المستمر بناءً على أهدافك المهنية.",
    "warning_fill_essential": "يرجى ملء جميع حقول المعلومات الأساسية قبل الحفظ.", # Simplified generic warning
    "warning_numeric_fields": "يرجى إدخال أرقام صحيحة لحقول الطلاب، الذكور، والإناث.",
    "success_pdf_generated": "تم إنشاء ملف الملاحظات PDF بنجاح.",
    "download_feedback_pdf": "📥 تنزيل ملف الملاحظات PDF",
    "checkbox_cleanup_sheets": "🪟 تنظيف أوراق LO غير المستخدمة (لا يوجد اسم مراقب)",
    "warning_sheets_removed": "تمت إزالة {} أوراق LO غير مستخدمة.",
    "info_reloaded_workbook": "تمت إعادة تحميل المصنف بعد التنظيف.",
    "info_no_sheets_to_cleanup": "لا يتم العثور على أوراق LO غير مستخدمة لتنظيفها.",
    "expander_guidelines": "📘 انقر هنا لعرض إرشادات الملاحظة",
    "info_no_guidelines": "ورقة الإرشادات فارغة أو تعذر قراءتها.",
    "warning_select_create_sheet": "يرجى تحديد أو إنشاء ورقة صالحة للمتابعة.",
    "label_overall_notes": "ملاحظات عامة لهذه الملاحظة الصفية",
}


# --- Function to get strings based on language toggle ---
def get_strings(arabic_mode):
    return ar_strings if arabic_mode else en_strings

# --- Function to determine performance level based on score ---
def get_performance_level(score, strings):
    # Ensure score is treated as a number for comparison, handle None/NaN/non-numeric
    if score is None or (isinstance(score, (int, float)) and math.isnan(score)) or not isinstance(score, (int, float)):
         return strings["overall_score_na"]

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
        # This catch should ideally not be hit if the initial check works, but included for safety
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
def generate_observation_pdf(data, feedback_content, strings): # Removed teacher_email as it's in data now
    buffer = io.BytesIO()
    # Note: For Arabic support in PDF, you need to register Arabic fonts with ReportLab
    # and potentially adjust directionality in styles. This requires more advanced ReportLab setup.
    # The current setup will likely render Arabic characters incorrectly or as boxes.
    doc = SimpleDocTemplate(buffer, pagesize=letter)

    story = []

    # --- Add School Logo ---
    school_name = data.get("school_name", "Default") # Use key from data dict
    logo_path = LOGO_PATHS.get(school_name, LOGO_PATHS["Default"])

    # Add logo only if file exists and is an image
    is_image_logo = False
    if os.path.exists(logo_path):
         try:
             # Basic check for image extension, not foolproof
             if logo_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                 img = Image(logo_path, width=1.5*inch, height=0.75*inch)
                 img.hAlign = 'CENTER'
                 story.append(img)
                 story.append(Spacer(1, 0.2*inch))
                 is_image_logo = True
             else:
                 # Not a recognized image format, add text title instead
                 story.append(Paragraph(strings["page_title"], styles['Heading1Centered']))
                 story.append(Spacer(1, 0.1*inch))
                 print(f"Logo file {logo_path} is not a standard image format. Using text title.")
         except Exception as e:
              print(f"Could not add logo for {school_name} from {logo_path}: {e}")
              story.append(Paragraph(f"[{school_name} Logo Placeholder]", styles['Normal'])) # Add placeholder text
              story.append(Spacer(1, 0.1*inch))
              # Fallback to text title if logo failed
              story.append(Paragraph(strings["page_title"], styles['Heading1Centered']))
              story.append(Spacer(1, 0.1*inch))
    else:
        print(f"Logo file not found for {school_name} at {logo_path}. Using text title.")
        story.append(Paragraph(strings["page_title"], styles['Heading1Centered']))
        story.append(Spacer(1, 0.1*inch))

    # Basic Information Table - Update keys to match data dictionary from input fields
    # Collect all basic info first to decide what to include
    basic_info_map = {
        strings["label_observer_name"]: data.get("observer_name", ""),
        strings["label_teacher_name"]: data.get("teacher_name", ""),
        strings["label_teacher_email"]: data.get("teacher_email", ""),
        strings["label_operator"]: data.get("operator", ""),
        strings["label_school_name"]: data.get("school_name", ""),
        strings["label_grade"]: data.get("grade", ""),
        strings["label_subject"]: data.get("subject", ""),
        strings["label_gender"]: data.get("gender", ""),
        strings["label_students"]: data.get("students", ""),
        strings["label_males"]: data.get("males", ""),
        strings["label_females"]: data.get("females", ""),
        strings["label_observation_date"]: data.get("observation_date", ""),
        strings["label_time_in"]: data.get("time_in", ""),
        strings["label_time_out"]: data.get("time_out", ""),
        # The duration label needs care with HTML/emoji
        strings["label_lesson_duration"].split("🕒")[0].strip(): data.get("duration_display", ""), # Pass formatted duration, strip emoji/html
        strings["label_period"]: data.get("period", ""),
        strings["label_obs_type"]: data.get("observation_type", ""),
    }

    # Format basic info data for the table, excluding None/empty values
    cleaned_basic_info_data = []
    # Ensure the Overall Score label and value are always last if score is available
    overall_score_pdf_row = None
    if data.get("overall_score_display") and data.get("overall_score_display") != strings["overall_score_na"]:
         overall_score_pdf_row = [strings["overall_score_label"] + ":", data.get("overall_score_display", strings["overall_score_na"])]


    for label, value in basic_info_map.items():
        # Only include if value is not None, not an empty string, and not NaN
        if value is not None and (not isinstance(value, str) or value.strip() != "") and not (isinstance(value, float) and math.isnan(value)):
             formatted_value = value
             # Format date/time objects nicely
             if isinstance(value, datetime):
                 formatted_value = value.strftime("%Y-%m-%d %H:%M")
             elif isinstance(value, date):
                 formatted_value = value.strftime("%Y-%m-%d")
             elif isinstance(value, time): # Use time object
                  formatted_value = value.strftime("%H:%M")
             elif isinstance(value, (int, float)):
                  formatted_value = str(value) # Convert numbers to string


             cleaned_basic_info_data.append([str(label) + ":", str(formatted_value)]) # Ensure keys are strings

    # Add the overall score row at the end if it exists
    if overall_score_pdf_row:
         cleaned_basic_info_data.append(overall_score_pdf_row)


    if cleaned_basic_info_data: # Only add table if there is data to display
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

         # Calculate optimal column widths: 30% for label, 70% for value
         available_width = letter[0] - (doc.leftMargin + doc.rightMargin)
         col_widths = [available_width * 0.3, available_width * 0.7]


         table = Table(cleaned_basic_info_data, colWidths=col_widths)
         table.setStyle(table_style)
         story.append(table)
         story.append(Spacer(1, 0.2*inch))


    # Rubric Scores - Needs to be adapted to the data structure collected from the inputs
    story.append(Paragraph(strings["subheader_rubric_scores"], styles['Heading2']))

    # Assuming 'data' dictionary passed to the PDF function contains domain_data
    domain_data = data.get("domain_data", {})

    if domain_data:
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
            if elements:
                 for element in elements:
                     label = element.get("label", "Unknown Element")
                     rating = element.get("rating", "N/A")
                     note = element.get("note", "")
                     descriptor = element.get("descriptor", "") # Specific descriptor text for the given rating

                     story.append(Paragraph(f"- <b>{label}:</b> Rating <b>{rating}</b>", styles['RubricElementRating']))

                     # Include Descriptor if available
                     if descriptor and isinstance(descriptor, str) and descriptor.strip():
                         # Clean and format the descriptor text - remove HTML/markdown like bold
                         cleaned_desc_para = re.sub(r'<.*?>', '', descriptor).replace('**', '')
                         desc_paragraphs = cleaned_desc_para.split('\n')
                         for desc_para in desc_paragraphs:
                             if desc_para.strip():
                                  story.append(Paragraph(desc_para, styles['RubricDescriptor']))
                         story.append(Spacer(1, 0.05*inch)) # Small space after descriptor

                     # Include Note if available
                     if note and isinstance(note, str) and note.strip():
                          # Clean and format the note text - remove HTML/markdown like bold
                          cleaned_note_para = re.sub(r'<.*?>', '', note).replace('**', '')
                          note_paragraphs = cleaned_note_para.split('\n')
                          story.append(Paragraph("  <i>Notes:</i>", styles['Normal'])) # Italicize notes header
                          for note_para in note_paragraphs:
                             if note_para.strip():
                                  story.append(Paragraph(note_para, styles['Normal']))
                          story.append(Spacer(1, 0.05*inch)) # Small space after notes

                     story.append(Spacer(1, 0.1*inch)) # Space after each element
            else:
                 story.append(Paragraph("No elements recorded for this domain.", styles['Normal']))


            story.append(Spacer(1, 0.2*inch)) # Space after each domain
    else:
        story.append(Paragraph("No rubric data available.", styles['Normal']))


    # Add Overall Notes
    overall_notes = data.get("overall_notes", "")
    if overall_notes and isinstance(overall_notes, str) and overall_notes.strip():
        story.append(Paragraph(strings["label_overall_notes"] + ":", styles['Heading2']))
        # Convert markdown in general notes if any
        cleaned_overall_notes = re.sub(r'<.*?>', '', overall_notes).replace('**', '<b>').replace('**', '</b>')
        notes_paragraphs = cleaned_overall_notes.split('\n')
        for note_para in notes_paragraphs:
            if note_para.strip():
                 story.append(Paragraph(note_para, styles['Normal']))
        story.append(Spacer(1, 0.2*inch))


    # Feedback Content (This part is crucial and needs to be generated dynamically)
    story.append(Paragraph("Feedback Report:", styles['Heading2']))
    if feedback_content and isinstance(feedback_content, str):
        # The feedback_content string needs to be constructed before calling this function.
        # It should include the greeting, summary of scores/judgments, performance summary,
        # suggested plan/steps, and closing.
        # Convert markdown-like text to ReportLab flowables
        feedback_paragraphs = feedback_content.split('\n\n') # Split by double newline
        for para in feedback_paragraphs:
            if para.strip():
                # Simple bold conversion and newline handling for ReportLab
                # Also clean any stray HTML tags
                para_styled = re.sub(r'<.*?>', '', para).replace('**', '<b>').replace('**', '</b>').replace('\n', '<br/>')
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

# --- Default File Definition ---
# Define the path to your template workbook
DEFAULT_FILE = "Teaching Rubric Tool_WeekTemplate.xlsx"


# Initialize workbook in session state if not already loaded
# This avoids reloading the workbook from disk on every single rerun
# BUT requires careful handling of saving modifications back to disk.
if 'workbook' not in st.session_state:
    st.session_state.workbook = None
    # Attempt to load the default workbook on first run
    if os.path.exists(DEFAULT_FILE):
        try:
            st.session_state.workbook = load_workbook(DEFAULT_FILE)
            st.success(strings["info_default_workbook"].format(DEFAULT_FILE))
        except Exception as e:
            st.error(strings["error_opening_default"].format(e))
            st.session_state.workbook = None
    else:
        st.warning(strings["warning_default_not_found"].format(DEFAULT_FILE))

# Use the workbook from session state
wb = st.session_state.workbook

# Sidebar page selection
page = st.sidebar.selectbox(strings["sidebar_select_page"], [strings["page_lesson_input"], strings["page_analytics"], strings["page_help"]]) # Added Help page

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

        # Email Domain Restriction
        # This widget manages st.session_state['auth_email_input'] automatically via its key.
        email = st.text_input("Enter your school email to continue", value=st.session_state.get('auth_email_input', ''), key='auth_email_input')

        allowed_domains = ["@charterschools.ae", "@adek.gov.ae"]
        # Check if email is entered AND if it ends with an allowed domain
        # Read the value directly from session state managed by the widget
        if not (st.session_state.get('auth_email_input') and any(st.session_state.get('auth_email_input', '').strip().lower().endswith(domain) for domain in allowed_domains)):
             if st.session_state.get('auth_email_input', '').strip(): # Only show specific warning if email is entered but invalid
                  st.warning("Access restricted. Please use an authorized school email.")
             st.stop() # Stop execution if email is invalid

        # REMOVED: st.session_state['auth_email_input'] = email # Store valid email - THIS LINE CAUSED THE ERROR


        lo_sheets = [sheet for sheet in wb.sheetnames if sheet.startswith("LO ")]
        if wb and lo_sheets:
             st.success(strings["success_lo_sheets_found"].format(len(lo_sheets)))

        # Cleanup unused LO sheets
        if wb and len(lo_sheets) > 1:
             if st.checkbox(strings.get("checkbox_cleanup_sheets", "🪟 Clean up unused LO sheets (no observer name)")):
                 to_remove = []
                 for sheet in lo_sheets:
                     if sheet == "LO 1":
                         continue
                     try:
                         aa1_value = wb[sheet]["AA1"].value
                         if aa1_value is None or (isinstance(aa1_value, str) and aa1_value.strip() == ""):
                             to_remove.append(sheet)
                     except Exception as e:
                         print(f"Could not check sheet '{sheet}' for cleanup due to error: {e}")
                         pass

                 if to_remove:
                     st.info(strings.get("warning_sheets_removed", "Removed {} unused LO sheets.").format(len(to_remove)))
                     for sheet in to_remove:
                         if sheet != "LO 1" and sheet in wb.sheetnames:
                             try:
                                 wb.remove(wb[sheet])
                             except Exception as e:
                                 st.error(f"Error removing sheet {sheet}: {e}")

                     try:
                          wb.save(DEFAULT_FILE)
                          st.success("Workbook saved after cleanup.")
                          st.session_state.workbook = load_workbook(DEFAULT_FILE)
                          wb = st.session_state.workbook
                          st.info(strings.get("info_reloaded_workbook", "Reloaded workbook after cleanup."))
                          st.rerun()
                     except Exception as e:
                          st.error(strings["error_saving_workbook"].format(e))
                 else:
                     st.info(strings.get("info_no_sheets_to_cleanup", "No unused LO sheets found to clean up."))


        # Display Guidelines expander (keeping it here for quick access)
        if wb and "Guidelines" in wb.sheetnames:
            guideline_content = []
            try:
                for row in wb["Guidelines"].iter_rows(values_only=True):
                    cleaned_row = [str(cell).strip() for cell in row if cell is not None]
                    guideline_content.extend(cleaned_row)
            except Exception as e:
                st.error(f"Error reading Guidelines sheet: {e}")
                guideline_content = [f"Error loading guidelines: {e}"]

            cleaned_guidelines = [line for line in guideline_content if line]
            if cleaned_guidelines:
                st.expander(strings.get("expander_guidelines", "📘 Click here to view observation guidelines")).markdown(
                    "\n".join(cleaned_guidelines)
                )
            else:
                st.info(strings.get("info_no_guidelines", "Guidelines sheet is empty or could not be read."))


        lo_sheets = [sheet for sheet in wb.sheetnames if sheet.startswith("LO ")]
        if "LO 1" not in wb.sheetnames:
            st.error(strings["error_template_not_found"])
            st.stop() # Cannot proceed without template

        existing_sheets_for_selection = sorted([s for s in lo_sheets if s != "LO 1"])
        sheet_selection_options = [strings["option_create_new"]] + existing_sheets_for_selection

        # --- Sheet Selection Dropdown ---
        selected_option = st.selectbox(
            strings["select_sheet_or_create"],
            sheet_selection_options,
            key='sheet_selector'
        )

        # --- Rubric Structure Definition ---
        rubric_domains_structure = {
            "Domain 1": ("I11", 5), # Starting Cell (Rating Column), Number of Elements
            "Domain 2": ("I20", 3),
            "Domain 3": ("I27", 4),
            "Domain 4": ("I35", 3),
            "Domain 5": ("I42", 2),
            "Domain 6": ("I48", 2),
            "Domain 7": ("I54", 2),
            "Domain 8": ("I60", 3),
            "Domain 9": ("I67", 2)
        }

        # --- Rubric Descriptor Reading ---
        rubric_descriptors = {}
        try:
             template_ws = wb["LO 1"]
             for domain, (start_cell, count) in rubric_domains_structure.items():
                  row_start = int(start_cell[1:])
                  for i in range(count):
                       row = row_start + i
                       element_key = f"{domain}_{i}"
                       descriptor_text_parts = []
                       try:
                            # Assuming descriptors are in columns E (index 4) to H (index 7)
                            for col_idx in range(4, 8):
                                 col_letter = get_column_letter(col_idx + 1)
                                 cell_value = template_ws[f"{col_letter}{row}"].value
                                 if cell_value is not None and isinstance(cell_value, str) and cell_value.strip():
                                      descriptor_text_parts.append(cell_value.strip())
                       except Exception as e:
                            print(f"Error reading descriptor cells for {domain} element {i}: {e}")
                            pass

                       rubric_descriptors[element_key] = {'general': " ".join(descriptor_text_parts) if descriptor_text_parts else strings["info_no_descriptors"]}

        except KeyError:
            st.warning("Template sheet 'LO 1' not found. Cannot read rubric descriptors.")
        except Exception as e:
            st.error(f"Error reading rubric descriptors from template: {e}")


        # --- Function to read existing data from a sheet (to pre-fill inputs) ---
        def load_existing_data(worksheet: Worksheet, rubric_structure):
            data = {}
            try: data["grade"] = worksheet["B1"].value
            except Exception: pass
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

            try:
                date_val = worksheet["D10"].value
                if isinstance(date_val, datetime):
                    data["observation_date"] = date_val.date()
                elif isinstance(date_val, date):
                    data["observation_date"] = date_val
                elif isinstance(date_val, str) and date_val:
                     try:
                          data["observation_date"] = datetime.strptime(date_val, "%Y-%m-%d").date()
                     except ValueError:
                          pass
            except Exception:
                pass

            try:
                time_in_val = worksheet["D7"].value
                if isinstance(time_in_val, datetime.time):
                    data["time_in"] = time_in_val
                elif isinstance(time_in_val, datetime):
                    data["time_in"] = time_in_val.time()
                elif isinstance(time_in_val, str) and time_in_val:
                    try:
                        data["time_in"] = datetime.strptime(time_in_val, "%H:%M:%S").time()
                    except ValueError:
                         try:
                              data["time_in"] = datetime.strptime(time_in_val, "%H:%M").time()
                         except ValueError:
                              pass
            except Exception:
                pass


            try:
                time_out_val = worksheet["D8"].value
                if isinstance(time_out_val, datetime.time):
                    data["time_out"] = time_out_val
                elif isinstance(time_out_val, datetime):
                    data["time_out"] = time_out_val.time()
                elif isinstance(time_out_val, str) and time_out_val:
                    try:
                        data["time_out"] = datetime.strptime(time_out_val, "%H:%M:%S").time()
                    except ValueError:
                         try:
                              data["time_out"] = datetime.strptime(time_out_val, "%H:%M").time()
                         except ValueError:
                              pass
            except Exception:
                pass

            try: data["period"] = worksheet["D4"].value
            except Exception: pass


            try: data["observer_name"] = worksheet["AA1"].value
            except Exception: pass
            try: data["teacher_name"] = worksheet["AA2"].value
            except Exception: pass
            try: data["observation_type"] = worksheet["AA3"].value
            except Exception: pass
            try: data["operator"] = worksheet["AA5"].value
            except Exception: pass
            try: data["school_name"] = worksheet["AA6"].value
            except Exception: pass
            try: data["overall_notes"] = worksheet["AA7"].value
            except Exception: pass
            try: data["teacher_email"] = worksheet["AA8"].value
            except Exception: pass
            try: data["feedback_generated_timestamp"] = worksheet["AA9"].value
            except Exception: pass


            data["element_inputs"] = {}
            for idx, (domain, (start_cell, count)) in enumerate(rubric_structure.items()):
                col_rating = start_cell[0]
                col_note = 'J'
                try:
                    row_start = int(start_cell[1:])
                    for i in range(count):
                        row = row_start + i
                        rating_key = f"{domain}_{i}_rating"
                        note_key = f"{domain}_{i}_note"

                        try:
                            rating_value_from_sheet = worksheet[f"{col_rating}{row}"].value
                            if isinstance(rating_value_from_sheet, float) and rating_value_from_sheet.is_integer():
                                rating_value_from_sheet = int(rating_value_from_sheet)
                            elif isinstance(rating_value_from_sheet, str) and rating_value_from_sheet.strip().upper() == "NA":
                                rating_value_from_sheet = "NA"
                            elif isinstance(rating_value_from_sheet, str) and rating_value_from_sheet.strip().isdigit():
                                rating_value_from_sheet = int(rating_value_from_sheet.strip())
                            elif rating_value_from_sheet is None or (isinstance(rating_value_from_sheet, str) and rating_value_from_sheet.strip() == ""):
                                rating_value_from_sheet = "NA"

                            data["element_inputs"][rating_key] = rating_value_from_sheet
                        except Exception:
                            data["element_inputs"][rating_key] = "NA"

                        try:
                            note_value_from_sheet = worksheet[f"{col_note}{row}"].value
                            data["element_inputs"][note_key] = note_value_from_sheet if note_value_from_sheet is not None else ""
                        except Exception:
                            data["element_inputs"][note_key] = ""

                except Exception as e:
                    print(f"Error loading rubric data for domain {domain} during load: {e}")
                    pass

            return data

        # --- State Management for Sheet Selection (Loading/Resetting) ---
        # Initialize tracker for the sheet whose data is currently loaded into session state
        if 'current_loaded_sheet_option' not in st.session_state:
             st.session_state['current_loaded_sheet_option'] = None # Initialize to no sheet loaded

        # Check if the selected sheet in the dropdown has changed from the one whose data is loaded
        if st.session_state['sheet_selector'] != st.session_state['current_loaded_sheet_option']:
             # The selected option has changed, update the tracker and load/reset state
             st.session_state['current_loaded_sheet_option'] = st.session_state['sheet_selector']

             if selected_option == strings["option_create_new"]:
                 # --- Logic for "Create New" ---
                 # Calculate the name for the *new* sheet (target name)
                 next_index = 1
                 existing_lo_numbers = []
                 for sheet in wb.sheetnames:
                      if sheet.startswith("LO "):
                           suffix = sheet[3:].strip()
                           if suffix.isdigit():
                                existing_lo_numbers.append(int(suffix))
                 if existing_lo_numbers:
                      next_index = max(existing_lo_numbers) + 1
                 calculated_new_sheet_name = f"LO {next_index}"

                 # Set the target sheet name for the form display and eventual saving
                 st.session_state['active_sheet_name_for_display'] = calculated_new_sheet_name
                 st.info(strings["subheader_filling_data"].format(st.session_state['active_sheet_name_for_display']))

                 # Reset session state keys to defaults for a new sheet
                 st.session_state.update({
                      'observer_name': '',
                      'teacher_name': '',
                      'teacher_email': '',
                      'operator': '',
                      'school_name': '',
                      'grade': '',
                      'subject': '',
                      'gender': '',
                      'students': 0,
                      'males': 0,
                      'females': 0,
                      'observation_date': datetime.now().date(),
                      'time_in': None,
                      'time_out': None,
                      'period': '',
                      'observation_type': strings["option_individual"],
                      'overall_notes': '',
                      'checkbox_send_feedback': False,
                      'element_inputs': {} # Initialize empty dict for element scores/notes
                 })

                 # Initialize element inputs in session state with default "NA" rating and empty notes
                 for idx, (domain, (start_cell, count)) in enumerate(rubric_domains_structure.items()):
                      for i in range(count):
                           rating_key = f"{domain}_{i}_rating"
                           note_key = f"{domain}_{i}_note"
                           st.session_state['element_inputs'][rating_key] = "NA"
                           st.session_state['element_inputs'][note_key] = ""

                 # Trigger a rerun to redraw the app with the newly reset state
                 st.rerun()

             else: # --- Logic for Selecting an Existing Sheet ---
                 target_sheet_name_from_selector = selected_option
                 st.session_state['active_sheet_name_for_display'] = target_sheet_name_from_selector
                 st.subheader(strings["subheader_filling_data"].format(st.session_state['active_sheet_name_for_display']))

                 try:
                     ws_to_load_from = wb[target_sheet_name_from_selector]

                     # Load existing data into session state
                     existing_data = load_existing_data(ws_to_load_from, rubric_domains_structure)

                     # Update session state with loaded data (excluding auth email)
                     for key, value in existing_data.items():
                          if key != 'auth_email_input':
                               # Handle specific types for widget defaults
                               if key in ['observer_name', 'teacher_name', 'teacher_email', 'operator', 'school_name', 'grade', 'subject', 'period', 'overall_notes']:
                                    st.session_state[key] = value if value is not None else '' # Use empty string for text inputs
                               elif key in ['students', 'males', 'females']:
                                    st.session_state[key] = value if value is not None else 0 # Use 0 for numbers
                               elif key in ['gender', 'observation_type']:
                                    # Ensure loaded value is one of the selectbox options or default
                                    options = ["Male", "Female", "Mixed", ""] if key == 'gender' else [strings["option_individual"], strings["option_joint"]]
                                    st.session_state[key] = value if value in options else (options[0] if options else '') # Default to first option or empty
                               elif key == 'observation_date':
                                    # Ensure it's a date object
                                    if isinstance(value, datetime):
                                         st.session_state[key] = value.date()
                                    elif isinstance(value, date):
                                         st.session_state[key] = value
                                    else:
                                         st.session_state[key] = datetime.now().date() # Default if invalid
                               elif key in ['time_in', 'time_out']:
                                     # Ensure it's a time object
                                     if isinstance(value, datetime.time):
                                          st.session_state[key] = value
                                     elif isinstance(value, datetime): # Sometimes loaded as datetime
                                          st.session_state[key] = value.time()
                                     else:
                                          st.session_state[key] = None # Default None
                               elif key == 'element_inputs':
                                     # element_inputs is a dictionary, merge it
                                     st.session_state[key] = value if isinstance(value, dict) else {}
                               else:
                                    st.session_state[key] = value # For other types


                     # Ensure element_inputs is initialized even if loading failed for it
                     if 'element_inputs' not in st.session_state or not isinstance(st.session_state['element_inputs'], dict):
                          st.session_state['element_inputs'] = {} # Safety initialization

                 except KeyError:
                     st.error(f"Error: Selected sheet '{target_sheet_name_from_selector}' not found or could not be accessed.")
                     st.session_state['sheet_selector'] = strings["option_create_new"] # Reset selector
                     st.session_state['current_loaded_sheet_option'] = strings["option_create_new"] # Reset tracker
                     st.session_state['active_sheet_name_for_display'] = None # Clear display name
                     st.rerun()
                     st.stop() # Stop current run
                 except Exception as e:
                     st.error(f"Error loading data from sheet '{target_sheet_name_from_selector}': {e}")
                     st.session_state['sheet_selector'] = strings["option_create_new"] # Reset selector
                     st.session_state['current_loaded_sheet_option'] = strings["option_create_new"] # Reset tracker
                     st.session_state['active_sheet_name_for_display'] = None # Clear display name
                     st.rerun()
                     st.stop() # Stop current run

                 # Trigger a rerun to redraw the app with the newly loaded state
                 st.rerun()

        # --- Input Form (Draw only if an active sheet name is set in state) ---
        # The form now reads its initial values from st.session_state,
        # which was populated by the state change logic above.
        if st.session_state.get('active_sheet_name_for_display'):
            st.info(f"**Currently Editing:** `{st.session_state['active_sheet_name_for_display']}`")

            # Basic Information Inputs
            st.markdown("---")
            st.subheader("Basic Information")
            cols = st.columns(2)
            with cols[0]:
                # Use keys ending in _form to isolate widget state from main session state keys
                st.text_input(strings["label_observer_name"], value=st.session_state.get('observer_name', '') or '', key='observer_name_input_form')
                st.text_input(strings["label_teacher_name"], value=st.session_state.get('teacher_name', '') or '', key='teacher_name_input_form')
                st.text_input(strings["label_teacher_email"], value=st.session_state.get('teacher_email', '') or '', key='teacher_email_input_form')
                st.text_input(strings["label_operator"], value=st.session_state.get('operator', '') or '', key='operator_input_form')
                st.text_input(strings["label_school_name"], value=st.session_state.get('school_name', '') or '', key='school_name_input_form')


            with cols[1]:
                st.text_input(strings["label_grade"], value=st.session_state.get('grade', '') or '', key='grade_input_form')
                st.text_input(strings["label_subject"], value=st.session_state.get('subject', '') or '', key='subject_input_form')
                gender_options = ["Male", "Female", "Mixed", ""]
                current_gender = st.session_state.get('gender', '') or ''
                gender_index = gender_options.index(current_gender) if current_gender in gender_options else len(gender_options) - 1
                st.selectbox(strings["label_gender"], gender_options, index=gender_index, key='gender_input_form')

                st.number_input(strings["label_students"], min_value=0, value=st.session_state.get('students', 0) or 0, step=1, key='students_input_form', format="%d")
                st.number_input(strings["label_males"], min_value=0, value=st.session_state.get('males', 0) or 0, step=1, key='males_input_form', format="%d")
                st.number_input(strings["label_females"], min_value=0, value=st.session_state.get('females', 0) or 0, step=1, key='females_input_form', format="%d")

            cols_date_time = st.columns(3)
            with cols_date_time[0]:
                 default_date = st.session_state.get('observation_date', datetime.now().date())
                 if isinstance(default_date, datetime): default_date = default_date.date()
                 elif not isinstance(default_date, date): default_date = datetime.now().date()
                 st.date_input(strings["label_observation_date"], value=default_date, key='observation_date_input_form')

            with cols_date_time[1]:
                default_time_in = st.session_state.get('time_in', None)
                if isinstance(default_time_in, datetime): default_time_in = default_time_in.time()
                st.time_input(strings["label_time_in"], value=default_time_in, key='time_in_input_form')

            with cols_date_time[2]:
                default_time_out = st.session_state.get('time_out', None)
                if isinstance(default_time_out, datetime): default_time_out = default_time_out.time()
                st.time_input(strings["label_time_out"], value=default_time_out, key='time_out_input_form')

            # Calculate and display Lesson Duration (reads from _form widget keys)
            lesson_duration_minutes = None
            duration_type = ""
            duration_display = strings["warning_calculate_duration"]

            time_in_val = st.session_state.get('time_in_input_form')
            time_out_val = st.session_state.get('time_out_input_form')

            if time_in_val is not None and time_out_val is not None:
                 try:
                      arbitrary_date = st.session_state.get('observation_date_input_form', date.today())
                      if isinstance(arbitrary_date, datetime): arbitrary_date = arbitrary_date.date()
                      elif not isinstance(arbitrary_date, date): arbitrary_date = date.today()

                      dt_in = datetime.combine(arbitrary_date, time_in_val)
                      dt_out = datetime.combine(arbitrary_date, time_out_val)

                      if dt_out < dt_in: dt_out += timedelta(days=1)

                      duration_seconds = (dt_out - dt_in).total_seconds()
                      lesson_duration_minutes = round(duration_seconds / 60)

                      if lesson_duration_minutes >= 40: duration_type = strings["duration_full_lesson"]
                      else: duration_type = strings["duration_walkthrough"]

                      duration_display = strings["label_lesson_duration"].format(lesson_duration_minutes, duration_type)

                 except Exception as e:
                      duration_display = strings["warning_could_not_calculate_duration"]
                      st.warning(f"Duration calculation error: {e}")

            st.info(duration_display)

            st.text_input(strings["label_period"], value=st.session_state.get('period', '') or '', key='period_input_form')
            obs_type_options = [strings["option_individual"], strings["option_joint"]]
            current_obs_type = st.session_state.get('observation_type', strings["option_individual"]) or strings["option_individual"]
            obs_type_index = obs_type_options.index(current_obs_type) if current_obs_type in obs_type_options else 0
            st.selectbox(strings["label_obs_type"], obs_type_options, index=obs_type_index, key='observation_type_input_form')

            st.markdown("---")
            st.subheader(strings["subheader_rubric_scores"])

            rating_options = ["NA", 1, 2, 3, 4, 5, 6]

            template_ws = None
            try: template_ws = wb["LO 1"]
            except KeyError: st.warning("Template sheet 'LO 1' not found. Cannot display rubric details."); template_ws = None

            if template_ws and rubric_domains_structure:
                for idx, (domain, (start_cell, count)) in enumerate(rubric_domains_structure.items()):
                     domain_title = domain
                     try:
                         title_row = int(start_cell[1:]) - 1
                         if f'B{title_row}' in template_ws and template_ws[f'B{title_row}'].value is not None:
                             domain_title = template_ws[f'B{title_row}'].value
                     except Exception: pass

                     st.markdown(f"#### {domain}: {domain_title}")

                     for i in range(count):
                          row = int(start_cell[1:]) + i
                          element_label = f"Element {i+1}"
                          try:
                               if f"B{row}" in template_ws and template_ws[f"B{row}"].value is not None:
                                    element_label = template_ws[f"B{row}"].value
                          except Exception: pass

                          element_key = f"{domain}_{i}"

                          st.markdown(f"##### {element_label}")

                          descriptor_text = rubric_descriptors.get(element_key, {}).get('general', strings["info_no_descriptors"])
                          st.expander(strings["expander_rubric_descriptors"]).markdown(descriptor_text)

                          cols_rating_note = st.columns(2)
                          with cols_rating_note[0]:
                             rating_key = f"{domain}_{i}_rating"
                             current_rating = st.session_state['element_inputs'].get(rating_key, "NA")
                             if isinstance(current_rating, float) and current_rating.is_integer(): current_rating = int(current_rating)
                             if current_rating not in rating_options: current_rating = "NA"

                             st.selectbox(
                                 strings["label_rating_for"].format(element_label),
                                 rating_options,
                                 index=rating_options.index(current_rating),
                                 key=f'{rating_key}_form' # Widget key
                             )
                          with cols_rating_note[1]:
                             note_key = f"{domain}_{i}_note"
                             current_note = st.session_state['element_inputs'].get(note_key, "") or ""
                             st.text_area(
                                 strings["label_write_notes"].format(element_label),
                                 value=current_note,
                                 key=f'{note_key}_form' # Widget key
                             )

            else:
                 st.info("Rubric structure could not be loaded from the template sheet.")


            st.markdown("---")
            st.text_area(strings["label_overall_notes"], value=st.session_state.get('overall_notes', '') or '', key='overall_notes_input_form')

            st.markdown("---")
            st.checkbox(strings["checkbox_send_feedback"], value=st.session_state.get('checkbox_send_feedback', False), key='send_feedback_checkbox_form')


            # --- Save Button ---
            if st.button(strings["button_save_observation"], key='save_observation_button'):

                 # --- Read Final Values from Form Widgets into Main Session State ---
                 # This step is crucial to get the latest values from the form widgets
                 # before performing validation and saving.
                 try:
                     st.session_state['observer_name'] = st.session_state.get('observer_name_input_form', '')
                     st.session_state['teacher_name'] = st.session_state.get('teacher_name_input_form', '')
                     st.session_state['teacher_email'] = st.session_state.get('teacher_email_input_form', '')
                     st.session_state['operator'] = st.session_state.get('operator_input_form', '')
                     st.session_state['school_name'] = st.session_state.get('school_name_input_form', '')
                     st.session_state['grade'] = st.session_state.get('grade_input_form', '')
                     st.session_state['subject'] = st.session_state.get('subject_input_form', '')
                     st.session_state['gender'] = st.session_state.get('gender_input_form', '')
                     st.session_state['students'] = st.session_state.get('students_input_form')
                     st.session_state['males'] = st.session_state.get('males_input_form')
                     st.session_state['females'] = st.session_state.get('females_input_form')
                     st.session_state['observation_date'] = st.session_state.get('observation_date_input_form')
                     st.session_state['time_in'] = st.session_state.get('time_in_input_form')
                     st.session_state['time_out'] = st.session_state.get('time_out_input_form')
                     st.session_state['period'] = st.session_state.get('period_input_form', '')
                     st.session_state['observation_type'] = st.session_state.get('observation_type_input_form', strings["option_individual"])
                     st.session_state['checkbox_send_feedback'] = st.session_state.get('send_feedback_checkbox_form', False)
                     st.session_state['overall_notes'] = st.session_state.get('overall_notes_input_form', '')


                     updated_element_inputs = {}
                     for domain, (start_cell, count) in rubric_domains_structure.items():
                          for i in range(count):
                               rating_key = f"{domain}_{i}_rating"
                               note_key = f"{domain}_{i}_note"
                               updated_element_inputs[rating_key] = st.session_state.get(f'{rating_key}_form', "NA")
                               updated_element_inputs[note_key] = st.session_state.get(f'{note_key}_form', "")
                     st.session_state['element_inputs'] = updated_element_inputs


                 except Exception as e:
                      st.error(f"Error reading form values before saving: {e}")
                      st.stop()


                 # --- Validation ---
                 essential_fields = {
                     strings["label_observer_name"]: st.session_state.get('observer_name'),
                     strings["label_teacher_name"]: st.session_state.get('teacher_name'),
                     strings["label_school_name"]: st.session_state.get('school_name'),
                     strings["label_grade"]: st.session_state.get('grade'),
                     strings["label_subject"]: st.session_state.get('subject'),
                     strings["label_gender"]: st.session_state.get('gender'),
                     strings["label_observation_date"]: st.session_state.get('observation_date'),
                 }
                 missing_essential = [label for label, value in essential_fields.items() if value is None or (isinstance(value, str) and value.strip() == "")]

                 if missing_essential:
                     st.warning(strings["warning_fill_essential"])
                     st.stop()

                 numeric_fields = {
                      strings["label_students"]: st.session_state.get('students'),
                      strings["label_males"]: st.session_state.get('males'),
                      strings["label_females"]: st.session_state.get('females'),
                 }
                 invalid_numeric = [label for label, value in numeric_fields.items() if value is not None and not isinstance(value, (int, float))]

                 if invalid_numeric:
                      st.warning(strings["warning_numeric_fields"])
                      st.stop()

                 all_ratings = [st.session_state['element_inputs'].get(f"{domain}_{i}_rating") for domain, (start_cell, count) in rubric_domains_structure.items() for i in range(count)]
                 if all(r is None or (isinstance(r, str) and r.upper() == "NA") for r in all_ratings):
                     st.warning("Please enter at least one rubric rating.")
                     st.stop()

                 # --- Get/Create Target Sheet ---
                 target_sheet_name = st.session_state['active_sheet_name_for_display']
                 ws = None
                 is_new_sheet = (st.session_state['sheet_selector'] == strings["option_create_new"])


                 if is_new_sheet:
                     try:
                         template_ws = wb["LO 1"]
                         ws = wb.copy_worksheet(template_ws)
                         ws.title = target_sheet_name
                         # After creating, change the selectbox state to the new sheet's name
                         # This will trigger a rerun and reload the newly created sheet's (empty) data
                         st.session_state['sheet_selector'] = target_sheet_name
                         # We don't need to rerun explicitly here, saving the workbook and reloading below will trigger it.
                         # However, we need to update the loaded sheet tracker to match the new selector value immediately
                         st.session_state['current_loaded_sheet_option'] = target_sheet_name

                         st.success(strings["success_sheet_created"].format(target_sheet_name))
                     except KeyError:
                         st.error(strings["error_template_not_found"])
                         st.stop()
                     except Exception as e:
                         st.error(f"Error creating new sheet '{target_sheet_name}': {e}")
                         st.stop()
                 else:
                     try:
                         ws = wb[target_sheet_name]
                     except KeyError:
                          st.error(f"Error: Target sheet '{target_sheet_name}' not found during save.")
                          st.stop()


                 # --- Write Data to Worksheet ---
                 try:
                     ws["AA1"].value = st.session_state.get('observer_name')
                     ws["AA2"].value = st.session_state.get('teacher_name')
                     ws["AA3"].value = st.session_state.get('observation_type')
                     ws["AA4"].value = datetime.now()
                     ws["AA5"].value = st.session_state.get('operator')
                     ws["AA6"].value = st.session_state.get('school_name')
                     ws["AA7"].value = st.session_state.get('overall_notes')
                     ws["AA8"].value = st.session_state.get('teacher_email')
                     # AA9 for feedback timestamp is updated later if PDF is generated

                     ws["B1"].value = st.session_state.get('grade')
                     ws["B5"].value = st.session_state.get('gender')
                     ws["B6"].value = st.session_state.get('students')
                     ws["B7"].value = st.session_state.get('males')
                     ws["B8"].value = st.session_state.get('females')
                     ws["D2"].value = st.session_state.get('subject')
                     ws["D4"].value = st.session_state.get('period')
                     ws["D7"].value = st.session_state.get('time_in')
                     ws["D8"].value = st.session_state.get('time_out')
                     obs_date = st.session_state.get('observation_date')
                     if obs_date and isinstance(obs_date, (date, datetime)):
                          ws["D10"].value = obs_date
                     else:
                          ws["D10"].value = None


                 except Exception as e:
                      st.error(f"Error writing basic information to sheet: {e}")


                 # Rubric Scores and Notes - Write values from session state
                 try:
                     for domain, (start_cell, count) in rubric_domains_structure.items():
                         col_rating = start_cell[0]
                         col_note = 'J'
                         row_start = int(start_cell[1:])
                         for i in range(count):
                             row = row_start + i
                             rating_key = f"{domain}_{i}_rating"
                             note_key = f"{domain}_{i}_note"

                             rating_value = st.session_state['element_inputs'].get(rating_key, "NA")
                             ws[f"{col_rating}{row}"].value = rating_value if rating_value != "NA" else "NA"

                             note_value = st.session_state['element_inputs'].get(note_key, "")
                             ws[f"{col_note}{row}"].value = note_value if note_value is not None else ""

                 except Exception as e:
                     st.error(f"Error writing rubric data to sheet: {e}")


                 # --- Calculate Overall Score and Domain Averages ---
                 domain_data_for_feedback = {}
                 overall_scores_list = []

                 template_ws = None
                 try: template_ws = wb["LO 1"]
                 except KeyError: pass


                 for domain_name, (start_cell, count) in rubric_domains_structure.items():
                      domain_ratings = []
                      elements_data = []

                      domain_title = domain_name
                      if template_ws:
                           try:
                                title_row = int(start_cell[1:]) - 1
                                if f'B{title_row}' in template_ws and template_ws[f'B{title_row}'].value is not None:
                                     domain_title = template_ws[f'B{title_row}'].value
                           except Exception: pass


                      for i in range(count):
                          rating_key = f"{domain_name}_{i}_rating"
                          note_key = f"{domain_name}_{i}_note"

                          element_label = f"Element {i+1}"
                          if template_ws:
                              try:
                                  label_row = int(start_cell[1:]) + i
                                  if f"B{label_row}" in template_ws and template_ws[f"B{label_row}"].value is not None:
                                       element_label = template_ws[f"B{label_row}"].value
                              except Exception: pass


                          rating = st.session_state['element_inputs'].get(rating_key)
                          note = st.session_state['element_inputs'].get(note_key, "")

                          if isinstance(rating, (int, float)) and not np.isnan(rating):
                               domain_ratings.append(float(rating))
                               overall_scores_list.append(float(rating))

                          element_key_for_descriptor = f"{domain_name}_{i}"
                          descriptor_text = rubric_descriptors.get(element_key_for_descriptor, {}).get('general', strings["info_no_descriptors"])

                          elements_data.append({
                             "label": element_label,
                             "rating": rating if rating is not None else "NA",
                             "note": note,
                             "descriptor": descriptor_text
                          })


                      domain_average = statistics.mean(domain_ratings) if domain_ratings else np.nan
                      domain_judgment = get_performance_level(domain_average, strings)

                      domain_data_for_feedback[domain_name] = {
                          "title": domain_title,
                          "average": domain_average,
                          "judgment": domain_judgment,
                          "elements": elements_data
                      }

                 overall_average_score = statistics.mean(overall_scores_list) if overall_scores_list else np.nan
                 overall_judgment = get_performance_level(overall_average_score, strings)

                 # --- Update Feedback Log Sheet ---
                 feedback_log_sheet_name = strings["feedback_log_sheet_name"]
                 try:
                     if feedback_log_sheet_name not in wb.sheetnames:
                         log_ws = wb.create_sheet(feedback_log_sheet_name)
                         log_ws.append(strings["feedback_log_header"])
                     else:
                         log_ws = wb[feedback_log_sheet_name]

                         existing_row_index = None
                         for row_idx, row in enumerate(log_ws.iter_rows(min_row=2, values_only=True)):
                              if row and row[0] == target_sheet_name:
                                   existing_row_index = row_idx + 2
                                   break

                         log_data_row = [
                             target_sheet_name,
                             st.session_state.get('observer_name'),
                             st.session_state.get('teacher_name'),
                             st.session_state.get('teacher_email'),
                             st.session_state.get('school_name'),
                             st.session_state.get('subject'),
                             st.session_state.get('observation_date'),
                             overall_judgment,
                             overall_average_score if not np.isnan(overall_average_score) else None,
                             st.session_state.get('overall_notes')
                         ]

                         if existing_row_index:
                              for col_idx, value in enumerate(log_data_row):
                                   log_ws.cell(row=existing_row_index, column=col_idx + 1, value=value)
                              st.success(f"Feedback log updated for sheet: {target_sheet_name}")
                         else:
                              log_ws.append(log_data_row)
                              st.success(strings["success_feedback_log_updated"])

                     st.session_state['feedback_generated_timestamp'] = datetime.now()


                 except Exception as e:
                     st.error(strings["error_updating_log"].format(e))


                 # --- Generate Feedback Report Text ---
                 feedback_content_text = ""
                 if st.session_state.get('send_feedback_checkbox_form', False):
                     try:
                         teacher_name = st.session_state.get('teacher_name', 'Teacher')
                         obs_date_display = st.session_state.get('observation_date')
                         if isinstance(obs_date_display, (date, datetime)): obs_date_display = obs_date_display.strftime("%Y-%m-%d")
                         else: obs_date_display = "N/A"

                         feedback_content_text += strings["feedback_greeting"].format(teacher_name, obs_date_display)
                         feedback_content_text += strings["feedback_observer"].format(st.session_state.get('observer_name', 'N/A'))
                         feedback_content_text += strings["feedback_duration"].format(f"{lesson_duration_minutes} minutes ({duration_type})" if lesson_duration_minutes is not None else 'N/A')
                         feedback_content_text += strings["feedback_subject_fb"].format(st.session_state.get('subject', 'N/A'))
                         feedback_content_text += strings["feedback_school"].format(st.session_state.get('school_name', 'N/A'))

                         feedback_content_text += strings["feedback_summary_header"]

                         if domain_data_for_feedback:
                              for domain_name, domain_info in domain_data_for_feedback.items():
                                   feedback_content_text += strings["feedback_domain_header"].format(domain_name, domain_info.get("title", domain_name))
                                   for element in domain_info.get("elements", []):
                                         feedback_content_text += strings["feedback_element_rating"].format(element["label"], element["rating"])
                                         if element["descriptor"] and element["descriptor"] != strings["info_no_descriptors"]:
                                               feedback_content_text += strings["feedback_descriptor_for_rating"].format(element["rating"], element["descriptor"])
                                         if element["note"] and element["note"].strip():
                                              feedback_content_text += f"  *Notes:* {element['note'].strip()}\n"

                                   if domain_info.get("average") is not None and not np.isnan(domain_info["average"]):
                                        feedback_content_text += strings["feedback_domain_average"].format(domain_info["average"])
                                   feedback_content_text += "\n"

                         if overall_average_score is not None and not np.isnan(overall_average_score):
                             feedback_content_text += strings["feedback_overall_score"].format(overall_average_score)

                         feedback_content_text += strings["feedback_performance_summary"]
                         feedback_content_text += strings["overall_performance_level_text"].format(overall_judgment) + "\n\n"

                         if overall_judgment == strings["perf_level_very_weak"]:
                              feedback_content_text += strings["feedback_support_plan_intro"] + strings["plan_very_weak_overall"] + "\n\n"
                         elif overall_judgment == strings["perf_level_weak"]:
                              feedback_content_text += strings["feedback_support_plan_intro"] + strings["plan_weak_overall"] + "\n\n"
                              for domain_name, domain_info in domain_data_for_feedback.items():
                                   if domain_info.get("judgment") == strings["perf_level_weak"]:
                                        feedback_content_text += strings["plan_weak_domain"].format(domain_info.get("title", domain_name), "...") + "\n"
                                        feedback_content_text += "[Provide specific actionable steps here]\n\n"

                         elif overall_judgment == strings["perf_level_acceptable"]:
                              feedback_content_text += strings["feedback_next_steps_intro"] + strings["steps_acceptable_overall"] + "\n\n"
                         elif overall_judgment == strings["perf_level_good"]:
                             feedback_content_text += strings["feedback_next_steps_intro"] + strings["steps_good_overall"] + "\n\n"
                             for domain_name, domain_info in domain_data_for_feedback.items():
                                   if domain_info.get("judgment") == strings["perf_level_good"]:
                                        feedback_content_text += strings["steps_good_domain"].format(domain_info.get("title", domain_name), "...") + "\n"
                                        feedback_content_text += "[Provide specific growth opportunities here]\n\n"

                         elif overall_judgment == strings["perf_level_excellent"]:
                              feedback_content_text += strings["feedback_next_steps_intro"] + strings["steps_excellent_overall"] + "\n\n"
                              for domain_name, domain_info in domain_data_for_feedback.items():
                                   if domain_info.get("judgment") == strings["perf_level_excellent"]:
                                        feedback_content_text += strings["steps_excellent_domain"].format(domain_info.get("title", domain_name), "...") + "\n"
                                        feedback_content_text += "[Provide opportunities for leadership/sharing here]\n\n"
                         else:
                             feedback_content_text += strings["no_specific_plan_needed"] + "\n\n"


                         feedback_content_text += strings["feedback_closing"]
                         feedback_content_text += strings["feedback_regards"]

                         st.session_state['generated_feedback_text'] = feedback_content_text
                         st.success(strings["success_feedback_generated"])

                     except Exception as e:
                          st.error(f"Error generating feedback text: {e}")
                          st.session_state['generated_feedback_text'] = "Error generating feedback text."
                          feedback_content_text = ""


                 # --- Generate PDF (if checkbox is checked) ---
                 pdf_buffer = None
                 if st.session_state.get('send_feedback_checkbox_form', False) and feedback_content_text:
                      st.info("Generating PDF feedback report...")
                      pdf_data = {
                          "observer_name": st.session_state.get('observer_name', ''),
                          "teacher_name": st.session_state.get('teacher_name', ''),
                          "teacher_email": st.session_state.get('teacher_email', ''),
                          "operator": st.session_state.get('operator', ''),
                          "school_name": st.session_state.get('school_name', ''),
                          "grade": st.session_state.get('grade', ''),
                          "subject": st.session_state.get('subject', ''),
                          "gender": st.session_state.get('gender', ''),
                          "students": st.session_state.get('students'),
                          "males": st.session_state.get('males'),
                          "females": st.session_state.get('females'),
                          "observation_date": st.session_state.get('observation_date'),
                          "time_in": st.session_state.get('time_in'),
                          "time_out": st.session_state.get('time_out'),
                          "duration_display": f"{lesson_duration_minutes} minutes ({duration_type})" if lesson_duration_minutes is not None else strings["warning_could_not_calculate_duration"],
                          "period": st.session_state.get('period', ''),
                          "observation_type": st.session_state.get('observation_type', strings["option_individual"]),
                          "overall_notes": st.session_state.get('overall_notes', ''),
                          "overall_score_display": strings["overall_score_value"].format(overall_average_score) if not np.isnan(overall_average_score) else strings["overall_score_na"],
                          "domain_data": domain_data_for_feedback,
                      }
                      pdf_buffer = generate_observation_pdf(pdf_data, feedback_content_text, strings)

                      if pdf_buffer:
                           st.success(strings["success_pdf_generated"])
                           try:
                                ws["AA9"].value = datetime.now()
                           except Exception as e:
                                st.warning(f"Could not write feedback timestamp to sheet {target_sheet_name}: {e}")

                      else:
                           st.error("Failed to generate PDF feedback report.")


                 # --- Save Workbook to File ---
                 try:
                      wb.save(DEFAULT_FILE)
                      st.success(strings["success_data_saved"])
                      st.session_state.workbook = load_workbook(DEFAULT_FILE)
                      # No explicit rerun needed here, as changing sheet_selector above already triggers one.
                      # However, explicitly updating wb seems safer immediately.
                      wb = st.session_state.workbook


                 except Exception as e:
                      st.error(strings["error_saving_workbook"].format(e))


                 # --- Provide Download Buttons ---
                 st.markdown("---")
                 st.markdown("##### Download Options")
                 col1, col2, col3 = st.columns(3)

                 with col1:
                     workbook_buffer = io.BytesIO()
                     # Use the potentially newly reloaded wb
                     st.session_state.workbook.save(workbook_buffer)
                     workbook_buffer.seek(0)
                     st.download_button(
                         label=strings["download_workbook"],
                         data=workbook_buffer,
                         file_name=DEFAULT_FILE,
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                     )

                 with col2:
                     if pdf_buffer:
                         teacher_name_for_pdf = st.session_state.get('teacher_name', 'feedback').replace(' ', '_').lower() or 'feedback'
                         obs_date_for_pdf = st.session_state.get('observation_date')
                         date_str_for_pdf = obs_date_for_pdf.strftime("%Y%m%d") if isinstance(obs_date_for_pdf, (date, datetime)) else "undated"
                         pdf_file_name = f"observation_feedback_{teacher_name_for_pdf}_{date_str_for_pdf}.pdf"

                         st.download_button(
                             label=strings["download_feedback_pdf"],
                             data=pdf_buffer,
                             file_name=pdf_file_name,
                             mime="application/pdf"
                         )
                     else:
                          st.button(strings["download_feedback_pdf"], disabled=True, help="Generate feedback first")

                 with col3:
                     feedback_log_sheet_name = strings["feedback_log_sheet_name"]
                     if feedback_log_sheet_name in st.session_state.workbook.sheetnames: # Use reloaded wb
                         try:
                             log_ws = st.session_state.workbook[feedback_log_sheet_name]
                             log_data = []
                             for row in log_ws.iter_rows(values_only=True):
                                  log_data.append(row)

                             if log_data:
                                 csv_buffer_log = io.StringIO()
                                 writer = csv.writer(csv_buffer_log)
                                 writer.writerows(log_data)
                                 csv_buffer_log.seek(0)

                                 st.download_button(
                                     label=strings["download_feedback_log_csv"],
                                     data=csv_buffer_log.getvalue(),
                                     file_name="feedback_log.csv",
                                     mime="text/csv"
                                 )
                             else:
                                  st.button(strings["download_feedback_log_csv"], disabled=True, help="Feedback log is empty.")
                         except Exception as e:
                             st.error(strings["error_generating_log_csv"].format(e))
                             st.button(strings["download_feedback_log_csv"], disabled=True, help="Error generating log CSV.")
                     else:
                         st.button(strings["download_feedback_log_csv"], disabled=True, help="Feedback Log sheet not found.")


                 if st.session_state.get('send_feedback_checkbox_form', False) and st.session_state.get('generated_feedback_text'):
                      st.markdown("---")
                      st.subheader("Generated Feedback Text (for verification)")
                      st.text(st.session_state['generated_feedback_text'])


        else:
             st.info(strings["warning_select_create_sheet"])


    # <--- End of Lesson Observation Input Page Logic --->


    elif page == strings["page_analytics"]:
        st.title(strings["title_analytics"])

        if not wb:
            st.warning(strings["warning_no_lo_sheets_analytics"])
            st.stop()

        lo_sheets_data_list = []
        feedback_log_data = pd.DataFrame()
        feedback_log_sheet_name = strings["feedback_log_sheet_name"]

        rubric_domains_avg_cells = {
            "Avg Domain 1": "I16", "Avg Domain 2": "I23", "Avg Domain 3": "I31",
            "Avg Domain 4": "I38", "Avg Domain 5": "I44", "Avg Domain 6": "I50",
            "Avg Domain 7": "I56", "Avg Domain 8": "I63", "Avg Domain 9": "I69"
        }

        try:
            lo_sheets_to_process = [sheet for sheet in wb.sheetnames if sheet.startswith("LO ") and sheet != "LO 1"]

            if not lo_sheets_to_process:
                 st.info(strings["warning_no_lo_sheets_analytics"])
            else:
                 for sheet_name in lo_sheets_to_process:
                     try:
                         ws = wb[sheet_name]
                         data = {
                             "Sheet": sheet_name,
                             "Observer": ws["AA1"].value if "AA1" in ws and ws["AA1"].value is not None else None,
                             "Teacher": ws["AA2"].value if "AA2" in ws and ws["AA2"].value is not None else None,
                             "Operator": ws["AA5"].value if "AA5" in ws and ws["AA5"].value is not None else None,
                             "School": ws["AA6"].value if "AA6" in ws and ws["AA6"].value is not None else None,
                             "Grade": ws["B1"].value if "B1" in ws and ws["B1"].value is not None else None,
                             "Subject": ws["D2"].value if "D2" in ws and ws["D2"].value is not None else None,
                             "Gender": ws["B5"].value if "B5" in ws and ws["B5"].value is not None else None,
                             "Students": pd.to_numeric(ws["B6"].value if "B6" in ws else None, errors='coerce'),
                             "Males": pd.to_numeric(ws["B7"].value if "B7" in ws else None, errors='coerce'),
                             "Females": pd.to_numeric(ws["B8"].value if "B8" in ws else None, errors='coerce'),
                             "Observation Date": ws["D10"].value if "D10" in ws else None,
                             "Observation Type": ws["AA3"].value if "AA3" in ws and ws["AA3"].value is not None else None,
                             "Overall Score": None,
                             "Overall Judgment": None,
                         }

                         for domain_key, cell_ref in rubric_domains_avg_cells.items():
                             try:
                                 avg_value = ws[cell_ref].value if cell_ref in ws else None
                                 data[domain_key] = pd.to_numeric(avg_value, errors='coerce')
                             except Exception:
                                 data[domain_key] = pd.NA

                         try:
                             overall_score_excel = ws["AM1"].value if "AM1" in ws else None
                             if isinstance(overall_score_excel, (int, float)) and not np.isnan(overall_score_excel):
                                 data["Overall Score"] = float(overall_score_excel)
                                 data["Overall Judgment"] = get_performance_level(data["Overall Score"], strings)
                             else:
                                 data["Overall Score"] = None
                                 data["Overall Judgment"] = strings["overall_score_na"]
                         except Exception:
                             data["Overall Score"] = None
                             data["Overall Judgment"] = strings["overall_score_na"]

                         lo_sheets_data_list.append(data)

                     except Exception as e:
                         st.warning(f"Could not load data from sheet '{sheet_name}' for analytics: {e}")

            if feedback_log_sheet_name in wb.sheetnames:
                try:
                     log_ws = wb[feedback_log_sheet_name]
                     data_rows = list(log_ws.iter_rows(min_row=2, values_only=True))
                     headers = [cell.value for cell in log_ws[1]]

                     if headers and data_rows:
                          cleaned_data_rows = [row for row in data_rows if any(cell is not None and str(cell).strip() != "" for cell in row)]
                          if cleaned_data_rows:
                              feedback_log_data = pd.DataFrame(cleaned_data_rows, columns=headers)
                              if 'Overall Score' in feedback_log_data.columns:
                                   feedback_log_data['Overall Score'] = pd.to_numeric(feedback_log_data['Overall Score'], errors='coerce')
                              if 'Date' in feedback_log_data.columns:
                                   feedback_log_data['Date'] = pd.to_datetime(feedback_log_data['Date'], errors='coerce')
                except Exception as e:
                    st.error(f"Error loading {feedback_log_sheet_name}: {e}")

            all_obs_data = pd.DataFrame(lo_sheets_data_list)

        except Exception as e:
            st.error(strings["error_loading_analytics"].format(e))
            all_obs_data = pd.DataFrame()


        if not all_obs_data.empty:
             if 'Observation Date' in all_obs_data.columns:
                 all_obs_data['Observation Date'] = pd.to_datetime(all_obs_data['Observation Date'], errors='coerce')

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


             st.markdown("---")
             st.subheader(strings["subheader_filter_analyze"])

             all_operators = sorted(all_obs_data['Operator'].dropna().unique().tolist()) if 'Operator' in all_obs_data.columns else []
             all_schools = sorted(all_obs_data['School'].dropna().unique().tolist()) if 'School' in all_obs_data.columns else []
             all_grades = sorted(all_obs_data['Grade'].dropna().unique().tolist()) if 'Grade' in all_obs_data.columns else []
             all_subjects = sorted(all_obs_data['Subject'].dropna().unique().tolist()) if 'Subject' in all_obs_data.columns else []
             all_teachers = sorted(all_obs_data['Teacher'].dropna().unique().tolist()) if 'Teacher' in all_obs_data.columns else []
             all_observers = sorted(all_obs_data['Observer'].dropna().unique().tolist()) if 'Observer' in all_obs_data.columns else []


             filter_operator = st.selectbox(strings["filter_by_operator"], [strings["option_all"]] + all_operators)
             filter_school = st.selectbox(strings["filter_by_school"], [strings["option_all"]] + all_schools)
             filter_grade = st.selectbox(strings["filter_by_grade"], [strings["option_all"]] + all_grades)
             filter_subject = st.selectbox(strings["filter_by_subject"], [strings["option_all"]] + all_subjects)
             filter_teacher = st.selectbox(strings["filter_teacher"], [strings["option_all"]] + all_teachers)
             filter_observer = st.selectbox(strings["filter_by_observer_an"], [strings["option_all"]] + all_observers)


             st.markdown("##### Filter by Date")
             today = datetime.now().date()
             valid_dates = all_obs_data['Observation Date'].dropna() if 'Observation Date' in all_obs_data.columns else pd.Series(dtype='datetime64[ns]')

             min_date_data = valid_dates.min().date() if not valid_dates.empty and isinstance(valid_dates.min(), datetime) else today - timedelta(days=365)
             max_date_data = valid_dates.max().date() if not valid_dates.empty and isinstance(valid_dates.max(), datetime) else today + timedelta(days=7)

             min_date_input = min_date_data if isinstance(min_date_data, date) else today - timedelta(days=365)
             max_date_input = max_date_data if isinstance(max_date_data, date) else today + timedelta(days=7)

             default_start_date = max(min_date_input, today - timedelta(days=365))
             default_start_date = min(default_start_date, today)

             default_end_date = max(max_date_input, today)
             default_end_date = max(default_end_date, default_start_date)


             try:
                  start_date = st.date_input(strings["filter_start_date"], value=default_start_date, min_value=min_date_input, max_value=max_date_input)
             except Exception:
                  start_date = st.date_input(strings["filter_start_date"], value=today - timedelta(days=365))

             try:
                  end_date = st.date_input(strings["filter_end_date"], value=default_end_date, min_value=min_date_input, max_value=max_date_input)
             except Exception:
                  end_date = st.date_input(strings["filter_end_date"], value=today + timedelta(days=7))


             filtered_data = all_obs_data.copy()

             if filter_operator != strings["option_all"] and 'Operator' in filtered_data.columns:
                  filtered_data = filtered_data[filtered_data['Operator'] == filter_operator].copy()

             if filter_school != strings["option_all"] and 'School' in filtered_data.columns:
                  filtered_data = filtered_data[filtered_data['School'] == filter_school].copy()

             if filter_grade != strings["option_all"] and 'Grade' in filtered_data.columns:
                  filtered_data = filtered_data[filtered_data['Grade'] == filter_grade].copy()

             if filter_subject != strings["option_all"] and 'Subject' in filtered_data.columns:
                  filtered_data = filtered_data[filtered_data['Subject'] == filter_subject].copy()

             if filter_teacher != strings["option_all"] and 'Teacher' in filtered_data.columns:
                  filtered_data = filtered_data[filtered_data['Teacher'] == filter_teacher].copy()

             if filter_observer != strings["option_all"] and 'Observer' in filtered_data.columns:
                  filtered_data = filtered_data[filtered_data['Observer'] == filter_observer].copy()

             if 'Observation Date' in filtered_data.columns and not filtered_data['Observation Date'].isna().all():
                  filtered_data = filtered_data.dropna(subset=['Observation Date']).copy()
                  start_timestamp = pd.Timestamp(start_date)
                  end_timestamp = pd.Timestamp(end_date) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

                  filtered_data = filtered_data[(filtered_data['Observation Date'] >= start_timestamp) & (filtered_data['Observation Date'] <= end_timestamp)].copy()


             st.markdown("---")
             st.subheader(strings["subheader_avg_score_filtered"])

             if not filtered_data.empty:
                 if 'Overall Score' in filtered_data.columns and not filtered_data['Overall Score'].isna().all():
                      avg_filtered_score = filtered_data['Overall Score'].mean()
                      st.write(f"Average Overall Score for Filtered Data: {avg_filtered_score:.2f}")
                 else:
                      st.write("Average Overall Score for Filtered Data: N/A (No valid scores found)")


                 st.markdown("#### Overall Judgment Distribution (Filtered)")
                 if 'Overall Judgment' in filtered_data.columns and not filtered_data['Overall Judgment'].isna().all():
                      valid_judgments = filtered_data['Overall Judgment'].dropna()
                      if not valid_judgments.empty:
                          judgment_counts = valid_judgments.value_counts()
                          judgment_order = [strings["perf_level_very_weak"], strings["perf_level_weak"], strings["perf_level_acceptable"], strings["perf_level_good"], strings["perf_level_excellent"], strings["overall_score_na"]]
                          judgment_counts = judgment_counts.reindex(judgment_order, fill_value=0).drop(strings["overall_score_na"], errors='ignore')
                          if not judgment_counts.empty and judgment_counts.sum() > 0:
                               st.bar_chart(judgment_counts)
                          else:
                               st.info("No observations with valid judgments found in the filtered data to chart.")
                      else:
                           st.info("No valid overall judgments found in the filtered data to chart.")
                 else:
                      st.info("Overall Judgment data is not available or is invalid in the filtered dataset.")


                 st.markdown("#### Average Score by Domain (Filtered)")
                 domain_avg_columns = [col for col in filtered_data.columns if col.startswith('Avg Domain')]
                 if domain_avg_columns:
                      domain_avg_data = filtered_data[domain_avg_columns].mean().reset_index()
                      domain_avg_data.columns = ['Domain', 'Average Score']
                      domain_avg_data['Domain'] = domain_avg_data['Domain'].str.replace('Avg ', '')

                      if not domain_avg_data.empty and not domain_avg_data['Average Score'].isna().all():
                          domain_avg_data = domain_avg_data.set_index('Domain')
                          st.bar_chart(domain_avg_data)
                      else:
                          st.info("No valid domain average scores found in the filtered data to chart.")
                 else:
                      st.info("Domain average data is not available in the dataset.")


                 st.markdown("---")
                 st.markdown("##### Filtered Observation Data Table")
                 display_columns = ['Sheet', 'Observer', 'Teacher', 'Operator', 'School', 'Grade', 'Subject', 'Observation Date', 'Overall Score', 'Overall Judgment'] + domain_avg_columns
                 display_columns = [col for col in display_columns if col in filtered_data.columns]
                 st.dataframe(filtered_data[display_columns])

                 st.markdown("###### Download Filtered Data")
                 col_csv, col_excel = st.columns(2)
                 with col_csv:
                      csv_buffer_filtered = io.StringIO()
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
                 st.info(strings["info_no_observation_data_filtered"])


             st.markdown("#### Average Score per Domain (Across all observations)")
             domain_avg_columns_all = [col for col in all_obs_data.columns if col.startswith('Avg Domain')]
             if domain_avg_columns_all:
                  domain_avg_data_all = all_obs_data[domain_avg_columns_all].mean().reset_index()
                  domain_avg_data_all.columns = ['Domain', 'Average Score']
                  domain_avg_data_all['Domain'] = domain_avg_data_all['Domain'].str.replace('Avg ', '')

                  if not domain_avg_data_all.empty and not domain_avg_data_all['Average Score'].isna().all():
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
                       st.info(strings["info_no_numeric_scores_overall"])
             else:
                  st.info("Domain average data columns not found in the dataset.")


             st.markdown("---")
             st.subheader(strings["subheader_teacher_performance"])
             st.info(strings["info_select_teacher"])

             selected_teacher_for_trend = st.selectbox("Select Teacher for Trend Analysis", [None] + all_teachers, format_func=lambda x: x if x is not None else "Select a Teacher...")


             if selected_teacher_for_trend:
                  teacher_data_for_trend = filtered_data[filtered_data['Teacher'] == selected_teacher_for_trend].copy()

                  if not teacher_data_for_trend.empty:
                       if 'Overall Score' in teacher_data_for_trend.columns and not teacher_data_for_trend['Overall Score'].isna().all():
                           st.subheader(strings["subheader_teacher_overall_avg"].format(selected_teacher_for_trend))
                           avg_teacher_score_filtered = teacher_data_for_trend['Overall Score'].mean()
                           st.write(f"Average Overall Score (Filtered): {avg_teacher_score_filtered:.2f}")
                       else:
                           st.write(f"Average Overall Score for {selected_teacher_for_trend} (Filtered): N/A (No valid scores found)")


                       domain_avg_columns_teacher = [col for col in teacher_data_for_trend.columns if col.startswith('Avg Domain')]

                       if 'Observation Date' in teacher_data_for_trend.columns and not teacher_data_for_trend['Observation Date'].isna().all() and domain_avg_columns_teacher:
                            st.subheader(strings["subheader_teacher_domain_trend"].format(selected_teacher_for_trend))

                            trend_data = teacher_data_for_trend.sort_values(by='Observation Date').copy()
                            trend_columns = ['Observation Date'] + domain_avg_columns_teacher
                            trend_data = trend_data[trend_columns].dropna(subset=['Observation Date'])
                            trend_data = trend_data.dropna(axis=1, how='all')


                            if not trend_data.empty and len(trend_data) > 1 and trend_data.select_dtypes(include=np.number).columns.tolist():
                                 trend_data = trend_data.set_index('Observation Date')
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
                                           file_name=f"{selected_teacher_for_trend.replace(' ', '_').lower()}_trend_data.csv",
                                           mime="text/csv"
                                      )
                                 with col_excel_trend:
                                      excel_buffer_trend = io.BytesIO()
                                      trend_data.to_excel(excel_buffer_trend)
                                      excel_buffer_trend.seek(0)
                                      st.download_button(
                                           label="📥 Download Trend Data (Excel)",
                                           data=excel_buffer_trend.getvalue(),
                                           file_name=f"{selected_teacher_for_trend.replace(' ', '_').lower()}_trend_data.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                      )

                            elif len(trend_data) <= 1:
                                st.info("Need at least 2 observations with valid dates and domain scores for this teacher under current filters to show a trend.")
                            elif not trend_data.select_dtypes(include=np.number).columns.tolist():
                                st.info("No numeric domain scores found for this teacher under current filters to show a trend.")
                            else:
                                st.info("Could not generate trend chart for this teacher with the available data under current filters.")


                       elif 'Observation Date' not in teacher_data_for_trend.columns or teacher_data_for_trend['Observation Date'].isna().all():
                            st.info("Observation dates are missing or invalid for this teacher under current filters.")
                       elif not domain_avg_columns_teacher:
                            st.info("Domain average data is not available for trend analysis.")

                  else:
                       st.info(strings["info_no_obs_for_teacher"])
             else:
                  st.info("Select a teacher above to view their performance trend.")


        else:
            st.info(strings["info_no_observation_data_filtered"])


    # <--- End of Lesson Observation Analytics Page Logic --->


    elif page == strings["page_help"]: # New Help/Guidelines page
        st.title(strings["title_help"])

        st.markdown("---")
        st.markdown("### Using the Application")
        st.markdown("""
        This tool allows you to record lesson observations using a standardized rubric and analyze the collected data.

        **Lesson Observation Input Page:**
        - **Workbook Loading:** The app attempts to load the default Excel workbook (`Teaching Rubric Tool_WeekTemplate.xlsx`). If not found, you'll be prompted. Ensure this file is in the same directory as the script or accessible path.
        - **Sheet Selection:** Choose an existing observation sheet ("LO #") or select "Create new" to start a fresh observation. When creating new, the app finds the next available "LO" number and copies the template sheet.
        - **Cleanup:** The "Clean up unused LO sheets" checkbox allows you to remove sheets that start with "LO " but do not have an Observer Name filled in Cell AA1. This helps keep your workbook organized.
        - **Data Entry:** Fill in the teacher details, lesson information, and rubric scores and notes for each element. Use the rubric guidance expanders to understand the criteria for each element.
        - **Generate Feedback Report:** Check the box if you want a PDF feedback report to be generated upon saving.
        - **Save Observation:** Click "Save Observation" to write the data to the selected (or newly created) sheet in the Excel workbook. The app will perform basic validation. It also updates a 'Feedback Log' sheet with key information and calculates overall/domain scores.
        - **Downloads:** After saving, you can download the updated Excel workbook, the generated PDF feedback report (if selected), and a CSV export of the Feedback Log.

        **Lesson Observation Analytics Page:**
        - This page reads data from all sheets starting with "LO " (excluding the template) and the "Feedback Log" sheet in the loaded workbook.
        - It displays overall statistics and allows you to filter the data by various criteria (School, Grade, Subject, Operator, Teacher, Observer, and Date Range).
        - You can view charts showing overall judgment distribution and average scores per domain for the filtered data, as well as overall domain averages across all observations.
        - A data table of the filtered observations is shown, with options to download it as CSV or Excel.
        - The Teacher Performance Over Time section allows you to select a specific teacher from the filtered data and view their domain score trends over time.

        **App Information and Guidelines Page:**
        - This page provides general information about the app and displays the observation guidelines read directly from the 'Guidelines' sheet in the Excel workbook.

        """)

        st.markdown("---")
        st.markdown("### Observation Guidelines")
        if wb and "Guidelines" in wb.sheetnames:
            guideline_content = []
            try:
                for row in wb["Guidelines"].iter_rows(values_only=True):
                    cleaned_row = [str(cell).strip() for cell in row if cell is not None]
                    guideline_content.extend(cleaned_row)
            except Exception as e:
                st.error(f"Error reading Guidelines sheet: {e}")
                guideline_content = [f"Error loading guidelines: {e}"]

            cleaned_guidelines = [line for line in guideline_content if line]
            if cleaned_guidelines:
                st.markdown("\n".join(cleaned_guidelines))
            else:
                st.info(strings.get("info_no_guidelines", "Guidelines sheet is empty or could not be read."))
        else:
            st.warning("Guidelines sheet not found in the workbook.")


# <--- This 'if wb:' block ends here.
else: # If workbook could not be loaded at the very start
    st.error("Could not load the workbook. Please ensure 'Teaching Rubric Tool_WeekTemplate.xlsx' exists and is accessible.")
