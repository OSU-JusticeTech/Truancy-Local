from datetime import datetime

from PyQt6.QtWidgets import QMessageBox, QInputDialog

from constructor import Student

BASE_HEADINGS = ["Last Name", "First Name", "Student #", "Age", "Grade", "Custodian", "Address", "Phone/Email",
                 "Suspension Hours", "Outcome of Correspondence",
                 "Date Preliminary Letter Sent", "Date Mediation Letter Sent",	"Mediation Date/Time"]

def add_report_to_sheet(window):
    """Add Total Absences column from PDF data to Excel file with ID/name matching and color coding"""
    
    # Check if we have both PDF data and Excel file
    if not window.check_files_ready():
        if not window.students:
            QMessageBox.warning(window, "No PDF Data", "Please load a PDF file first")
            return

        if not window.workbook:
            QMessageBox.warning(window, "No Excel File", "Please open an Excel file first")
            return
    
    try:
        # Get the sheet matching the index in the dropdown
        sheet_idx = window.sheets_combo.currentIndex() - 1
        if sheet_idx < 0:
            # Make new sheet
            sheet = blank_sheet(window.workbook, window.school_name)
            window.check_files_ready(did_update=True) # Refresh dropdown with new sheet added
        else:
            sheet = window.workbook.sheets[window.sheets_combo.currentText()]


        # Set date based on selector
        label = window.date_select.date().toString("MM/dd/yyyy")


        # Ask user for confirmation about going through with adding data
        confirm = QMessageBox.question(window, "Adding data to sheet", f"Adding column {label} to\n{sheet.name}\nContinue?",
                                      QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
        if confirm == QMessageBox.StandardButton.Cancel:
            return

        # Find locations of all columns
        column_locs = {}
        extra_column_strings = []
        extra_column_nums = []
        for col in range(1, sheet.used_range.last_cell.column + 1):
            header_value = sheet.range((1, col)).value
            if header_value in BASE_HEADINGS:
                column_locs[header_value] = col
            else:
                extra_column_strings.append(f"{_col_letter(col)} | {header_value}")
                extra_column_nums.append(col)
        # If required header isn't found, ask user for its new location
        for header in BASE_HEADINGS:
            if header not in column_locs:
                chosen, ok = QInputDialog.getItem(window, "Missing Column",
                                                  f"Could not find column '{header}'.\nSelect column to use instead:",
                                                  extra_column_strings, editable=False)
                if not ok:
                    return

                column_locs[header] = extra_column_nums[extra_column_strings.index(chosen)]

        # Insert data before last column
        if "Outcome of Correspondence" in column_locs:
            insert_col = column_locs["Outcome of Correspondence"]  # Insert pushes existing column to the right
        else:
            insert_col = sheet.used_range.last_cell.column + 1  # Default to final column

        # Move to the sheet before adding
        window.go_to_cell(sheet.name, f'{_col_letter(insert_col)}1:{_col_letter(insert_col + 1)}1')
        # sheet.select()
        # sheet.range(f'{_col_letter(insert_col)}1').select()

        # Insert two new columns
        sheet.range(f'{_col_letter(insert_col)}:{_col_letter(insert_col)}').api.Insert()
        sheet.range(f'{_col_letter(insert_col + 1)}:{_col_letter(insert_col + 1)}').api.Insert()
        # Shift known column locations to compensate for new columns
        for heading in column_locs:
            if column_locs[heading] >= insert_col:
                column_locs[heading] += 2
 
        # Find the last row by counting up to the last non-clear row
        last_row = sheet.used_range.last_cell.row
        while not sheet.range((last_row, 1)).value:
            last_row -= 1

        # Add header with user's label
        header_cell_ex = sheet.range((1, insert_col))
        header_cell_ex.value = f"{label} Excused Absences"
        header_cell_unex = sheet.range((1, insert_col + 1))
        header_cell_unex.value = f"{label} Unexcused Absences"

        # Clear all colors from the new column
        for row in range(1, last_row + 1):
            sheet.range((row, insert_col)).color = None
            sheet.range((row, insert_col + 1)).color = None
        
        print(f" ADDING ABSENCES WITH MATCHING ")
        
        # Build lookup dictionaries from PDF students
        pdf_by_id = {}  # {student_id: student_object}
        pdf_by_name = {}  # {(last_name, first_name): student_object}
        unmatched = set()
        
        for student in window.students:
            # Add to ID lookup
            if student.id:
                pdf_by_id[str(student.id).strip()] = student
            
            # Add to name lookup
            if student.lastName and student.firstName:
                last = str(student.lastName).strip().lower()
                first = str(student.firstName).strip().lower()
                pdf_by_name[(last, first)] = student

            # Add to set tracking students that haven't been matched
            unmatched.add(student)
        
        print(f"PDF students indexed: {len(pdf_by_id)} by ID, {len(pdf_by_name)} by name")
        
        # Get last row in Excel
        print(f"Excel has {last_row - 1} data rows (rows 2-{last_row})")
        
        # Track data
        no_match = 0

        
        # Loop through Excel rows and match
        for row in range(2, last_row + 1):
            # Get Student ID from Excel
            excel_student_id = sheet.range((row, column_locs["Student #"])).value
            
            # Get names from Excel
            excel_first_name = str(sheet.range((row, column_locs["First Name"])).value)
            excel_last_name = str(sheet.range((row, column_locs["Last Name"])).value)
            
            matched_student = None
            
            # Match by Student ID
            if excel_student_id:
                excel_student_id_str = str(int(excel_student_id)).strip()
                if excel_student_id_str in pdf_by_id:
                    matched_student = pdf_by_id[excel_student_id_str]
                    print(f"Row {row}: Matched by ID {excel_student_id_str}")
            
            # Check if name matches just in case
            if matched_student and excel_last_name and excel_first_name:
                last = str(excel_last_name).strip().lower()
                first = str(excel_first_name).strip().lower()

                if (last, first) in pdf_by_name and pdf_by_name[(last, first)] != matched_student:
                    print(f"!!! Student name mismatch: ID {excel_student_id} matches {matched_student.firstName} "
                          f"{matched_student.lastName} in Excel, {first} {last} in PDF")
            
            # Write Unexcused if matches are found
            if matched_student:
                # Remove from unmatched set
                unmatched.remove(matched_student)

                add_student(sheet, matched_student, insert_col, row, column_locs["Suspension Hours"])

            else:
                # No match found; leave blank no value
                print(f"Row {row}: No match found for {excel_first_name} {excel_last_name} (ID: {excel_student_id})")
                no_match += 1

                # Add "no data" to the new entry
                sheet.range((row, insert_col)).value = "no data"
                sheet.range((row, insert_col + 1)).value = "no data"
                sheet.range((row, insert_col)).color = (217, 217, 217)
                sheet.range((row, insert_col + 1)).color = (217, 217, 217)

        # Suffixes to add to grade number. Matched by index in list
        grade_suffix = ["", "st", "nd", "rd", "th", "th", "th", "th", "th", "th", "th", "th", "th"]

        new_list = set()

        # Add new rows for unmatched students
        for student in unmatched:
            insert_row = 1
            cont = True
            while cont:
                insert_row += 1
                currname = sheet.range((insert_row, column_locs["Last Name"])).value
                cont = bool(currname) and student.lastName > str(currname)


            sheet.range(f"{insert_row}:{insert_row}").insert('down')
            sheet.range(f"{insert_row}:{insert_row}").color = None

            add_student(sheet, student, insert_col, insert_row, column_locs["Suspension Hours"])

            try:
                gradenum = int(student.grade)
                grade = str(gradenum) + grade_suffix[gradenum]
            except:
                grade = student.grade

            sheet.range((insert_row, column_locs["Last Name"])).value = student.lastName
            sheet.range((insert_row, column_locs["First Name"])).value = student.firstName
            sheet.range((insert_row, column_locs["Student #"])).value = student.id
            sheet.range((insert_row, column_locs["Age"])).value = student.age
            sheet.range((insert_row, column_locs["Grade"])).value = grade

            new_list.add(sheet.range((insert_row, column_locs["Student #"])).value)

            last_row += 1

        # Groups of students for report table
        groups = {-1: [],
                  0: [],
                  1: [],
                  2: [],
                  3: []}

        # Loop through students again for report
        for row in range(2, last_row + 1):
            track_group(groups, sheet, row, insert_col, column_locs, new_list)


        # Print summary
        print(f"\n=== SUMMARY ===")
        print(f"No match found: {no_match}")
        print(f"Total rows processed: {last_row - 1}")

        # Write results to status box
        window.status_box.report_update(groups, label, Student.redThreshold, sheet.name,
                                        (_col_letter(insert_col), _col_letter(insert_col + 1)))
        # Update checkbox
        window.step_containers[2].setTitle("3. ☑ (requires steps 1 and 2)")
        
    except Exception as e:
        import traceback
        print(f"Error adding absences: {e}")
        print(traceback.format_exc())
        QMessageBox.critical(window, "Error", f"Error adding absences: {e}")

def blank_sheet(workbook, name):
    # Create new sheet
    shortname = name if len(name) < 32 else name[:30] + "…"
    sheet = workbook.sheets.add(name=shortname)

    sheet.range('A2').value = ""
    # Add base headings to sheet
    for i, heading in enumerate(BASE_HEADINGS):
        sheet.range((1, i + 1)).value = heading
        sheet.range((1, i + 1)).font.bold = True

    suspension_idx = BASE_HEADINGS.index("Suspension Hours") + 1
    sheet.range((1, suspension_idx)).color = (255, 255, 0)

    return sheet

def add_student(sheet, student, column, row, suspension_col):

    if student.unexcused:
        try:
            excused = float(student.excused)
            unexcused = float(student.unexcused)
            suspension = float(student.suspension)
            total_no_suspension = float(student.absenceTotal) - suspension
            cell_ex = sheet.range((row, column))
            cell_ex.value = excused
            cell_unex = sheet.range((row, column + 1))
            cell_unex.value = unexcused

            # Check for mismatch with report's total and calculated total
            if abs(excused + unexcused - total_no_suspension) > 0.02:
                print(f"!!! Total hours mismatch for {student.firstName} {student.lastName}"
                      f": Excel says {excused + unexcused}, PDF says {total_no_suspension}")

            # Update suspensions column
            sheet.range((row, suspension_col)).value = suspension
            sheet.range((row, suspension_col)).color = (255, 255, 0)

            # Color code based on absence hours
            if unexcused + excused >= Student.redThreshold:
                # Red for over limit
                cell_ex.color = (255, 0, 0)  # Red
                cell_unex.color = (255, 0, 0)

        except (ValueError, TypeError):
            print(f"Warning: Could not convert an absence total")
            # Invalid data
            sheet.range((row, column)).value = "no data"
            sheet.range((row, column + 1)).value = "no data"
            sheet.range((row, column)).color = (217, 217, 217)
            sheet.range((row, column + 1)).color = (217, 217, 217)
    else:
        # Student matched but has no absence data; no color
        sheet.range((row, column)).value = "no data"
        sheet.range((row, column + 1)).value = "no data"
        sheet.range((row, column)).color = (217, 217, 217)
        sheet.range((row, column + 1)).color = (217, 217, 217)



def track_group(groups, sheet, row, column, column_locs, new_list):

    # Add strings from the letter status columns
    prelim = sheet.range((row, column_locs["Date Preliminary Letter Sent"])).value
    if prelim is None:
        prelim = ""
    elif isinstance(prelim, datetime):
        prelim = prelim.strftime('%m/%d/%Y')
    else:
        prelim = str(prelim)
    mediation = sheet.range((row, column_locs["Date Mediation Letter Sent"])).value
    if mediation is None:
        mediation = ""
    elif isinstance(mediation, datetime):
        mediation = mediation.strftime('%m/%d/%Y')
    else:
        mediation = str(mediation)

    name = f"{sheet.range((row, column_locs['Last Name'])).value}, {sheet.range((row, column_locs['First Name'])).value}"

    student_data = (name, row, prelim, mediation)


    history = [] # Last three weeks' status. True = over limit, False = under limit, None = no data

    # March thru previous weeks to record history of being over the limit
    for c in range(max(column_locs["Suspension Hours"] + 1, column-4), column+2, 2):
        val_ex = sheet.range((row, c)).value
        val_unex = sheet.range((row, c + 1)).value
        try:
            val_int = int(val_ex) + int(val_unex)
            history.append(val_int >= Student.redThreshold)
        except (TypeError, ValueError):
            history.append(None)

    # Add students to groups based on whether they were over limits the last three weeks
    updated_history = [False * (3 - len(history))] + [bool(x) for x in history]

    if sheet.range((row, column_locs["Student #"])).value in new_list:
        groups[0].append(student_data)

    count = 0
    while updated_history[-1 - count] and count < len(history):
        count += 1

    if count > 0:
        groups[count].append(student_data)
    elif updated_history[-2] and history[-1] is not None:
        groups[-1].append(student_data)



def _col_letter(col_num):
    """Convert column number to Excel column letter (1=A, 2=B, ...)"""
    string = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        string = chr(65 + remainder) + string
    return string
