["B5"] = gender
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
        ws["Z7"] = "Remarks"
        ws["AA7"] = remarks

        if "Observation Log" not in wb.sheetnames:
            log_ws = wb.create_sheet("Observation Log")
            log_ws.append(["Sheet", "Observer", "Teacher", "Operator", "School", "Type", "Timestamp", "Remarks"])
        else:
            log_ws: Worksheet = wb["Observation Log"]

        log_ws.append([sheet_name, observer, teacher, operator, school, obs_type, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), remarks])

        save_path = f"updated_{sheet_name}.xlsx"
        wb.save(save_path)
        with open(save_path, "rb") as f:
            st.download_button("📥 Download updated workbook", f, file_name=save_path)
        os.remove(save_path)













  

  










