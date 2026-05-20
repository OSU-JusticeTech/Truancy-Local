
from PyQt6.QtWidgets import QFileDialog
from PyQt6.QtCore import QSettings

from pdf_parser import extract_students_from_pdf
import subprocess
import os
from datetime import datetime


def select_pdf(window):

    saved_pdf_dir = window.settings.value("pdf_dir", "/home")

    pdf_path = QFileDialog.getOpenFileName(window, "Open Truancy Report", saved_pdf_dir, "PDF (*.pdf)")[0]
    if not pdf_path:
        return
    
    # Should be saved in registry
    window.settings.setValue("pdf_dir", os.path.dirname(pdf_path))
    window.settings.sync()
    window.pdf_path_bar.setText(pdf_path)
    window.pdf_path_bar.repaint() # Manual repaint so box is filled before PDF is parsed

    school_name, generated_date, students = extract_students_from_pdf(pdf_path)

    print("School Name:", school_name)

    if len(students) == 0:
        print("No students")
    #     QMessageBox.warning(self, "No Data", "No students found in PDF")

    else:
        students[0].printHeaders()
        for s in students:
            s.print()

    if generated_date is None:
        now = datetime.now().date()
        generated_date = (now.month, now.day, now.year)

    window.pdf_opened.emit(pdf_path, students, school_name, generated_date)


def open_pdf(window):
    # Open the PDF with system's default viewer
    subprocess.Popen([window.pdf_path], shell=True)