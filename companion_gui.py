import xlwings
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton, \
    QScrollArea, QLineEdit, QComboBox, QVBoxLayout, QGroupBox, QHBoxLayout, QDateEdit, QLayoutItem, \
    QDialog, QLabel, QMessageBox
from PyQt6.QtCore import pyqtSlot, pyqtSignal, QSettings, Qt, QDate
import xlwings as xw
from docxtpl import DocxTemplate

from companion_btns.open_pdf import select_pdf, open_pdf
from companion_btns.open_excel import open_excel
from companion_btns.open_docx import open_docx
from companion_btns.generate_template import generate_template
from companion_btns.add_report_to_sheet import add_report_to_sheet
from difflib import SequenceMatcher
import os
import subprocess

from companion_btns.status_box import StatusBox


class TruancyWindow(QMainWindow):

    pdf_opened = pyqtSignal(str, list, str, tuple)
    excel_opened = pyqtSignal(str)
    docx_opened = pyqtSignal(str, object)

    def __init__(self):
        super().__init__()

        self.setWindowTitle("TruancyRecorder")
        self.setMinimumWidth(375)
        self.settings = QSettings("TruancyApp", "TruancyRecorder")

        #self.settings.setValue("agreed_policy", 0)

        # Show popup at startup, internal use policy, unless already agreed
        if not self.settings.value("agreed_policy", 0):
            if not self.show_internal_use_policy():
                import sys
                sys.exit()
            else:
                self.settings.setValue("agreed_policy", 1)
        
        # Store loaded students, workbook, and letter
        self.pdf_path = ""
        self.students = []
        self.school_name = ""
        self.workbook = None
        self.docx_path = ""
        self.docx_template = None

        # Associated with select_pdf and open_pdf
        select_pdf_button = QPushButton("Select Report PDF")
        select_pdf_button.clicked.connect(lambda: select_pdf(self))
        select_pdf_button.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "assets/pdf.png")))
        self.open_pdf_button = QPushButton()
        self.open_pdf_button.clicked.connect(lambda: open_pdf(self))
        self.open_pdf_button.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "assets/open.png")))
        self.open_pdf_button.setFixedWidth(30)
        self.open_pdf_button.setFlat(True)
        self.pdf_opened.connect(self.update_students)
        self.pdf_path_bar = QLineEdit()
        self.pdf_path_bar.setReadOnly(True)
        self.pdf_path_bar.setPlaceholderText("No PDF loaded")

        # Associated with open excel
        excel_button = QPushButton("Connect to Excel")
        excel_button.clicked.connect(lambda: open_excel(self))
        excel_button.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "assets/excel.png")))
        self.excel_opened.connect(self.update_workbook)
        self.excel_path_bar = QLineEdit()
        self.excel_path_bar.setReadOnly(True)
        self.excel_path_bar.setPlaceholderText("No Excel file selected")

        # Associated with open_docx
        select_docx_button = QPushButton("Select Letter Template")
        select_docx_button.clicked.connect(lambda: open_docx(self))
        select_docx_button.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "assets/word.png")))
        self.docx_opened.connect(self.update_docx)
        self.docx_path_bar = QLineEdit()
        self.docx_path_bar.setReadOnly(True)
        self.docx_path_bar.setPlaceholderText("No Word document loaded")

        # Button to add report to sheet
        self.add_absences_button = QPushButton("Add Report to Sheet")
        self.add_absences_button.clicked.connect(lambda: add_report_to_sheet(self))
        self.sheets_combo = QComboBox()
        self.date_select = QDateEdit()
        #self.date_select.setMaximumWidth(80)

        # Button for DOCX generation
        self.generate_template_button = QPushButton("Generate Letter for Selected Row in Excel")
        self.generate_template_button.clicked.connect(lambda: generate_template(self))

        # Text box to hold status messages for user
        self.status_box = StatusBox()
        self.status_box.go_to_cell.connect(self.go_to_cell)
        status_scroll = QScrollArea()
        status_scroll.setWidget(self.status_box)
        status_scroll.setWidgetResizable(True)

        center_layout = QVBoxLayout()

        def contain_widgets(group_name, widgets):
            # Encapsulates widgets in a named box
            container = QGroupBox(group_name)
            hlayout = QHBoxLayout()
            for w in widgets:
                hlayout.addWidget(w)
            container.setLayout(hlayout)
            return container

        self.step_containers = [
            contain_widgets("1. ☐", [excel_button, self.excel_path_bar]),
            contain_widgets("2. ☐", [select_pdf_button, self.pdf_path_bar, self.open_pdf_button]),
            contain_widgets("3. ☐ (requires steps 1 and 2)", [self.add_absences_button, self.sheets_combo, self.date_select]),
            status_scroll,
            contain_widgets("4. ☐", [select_docx_button, self.docx_path_bar]),
            contain_widgets("5. (requires steps 1 and 4)", [self.generate_template_button]),
        ]
        for sc in self.step_containers:
            center_layout.addWidget(sc)
        
        center_widget = QWidget()
        center_widget.setLayout(center_layout)
        self.setCentralWidget(center_widget)

        # Keep window above all other windows
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)

        self.check_files_ready()

    def open_policy_file(self):
        """Opens the Internal Use Policy Word document"""
        terms_file = os.path.join(os.path.dirname(__file__), "assets", "InternalUsePolicy.docx")
        if not os.path.exists(terms_file):
            QMessageBox.warning(self, "File Not Found", "Could not find InternalUsePolicy.docx")
            return
        try:
            os.startfile(terms_file)  # Windows - opens with default app (Word)
        except AttributeError:
            subprocess.Popen(["xdg-open", terms_file])  # Linux
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open file: {e}")

    def show_internal_use_policy(self):
        """Shows policy popup with View Full Policy button, returns True if user accepts"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Internal Use Policy")
        dialog.setModal(True)

        terms_summary = QLabel(
            "TRUANCY RECORDER - INTERNAL USE POLICY\n\n"
            "This application processes student absence data for educational purposes only.\n\n"
            "By using this application, you agree to our full Internal Use Policy.\n\n"
            "Do you accept these terms?"
        )
        terms_summary.setWordWrap(True)
        terms_summary.setAlignment(Qt.AlignmentFlag.AlignCenter)

        view_terms_button = QPushButton("View Full Policy")
        view_terms_button.clicked.connect(self.open_policy_file)

        accept_button = QPushButton("I Accept")
        accept_button.clicked.connect(dialog.accept)

        decline_button = QPushButton("I Don't Accept")
        decline_button.clicked.connect(dialog.reject)

        button_layout = QHBoxLayout()
        button_layout.addWidget(accept_button)
        button_layout.addWidget(decline_button)

        layout = QVBoxLayout()
        layout.addWidget(terms_summary)
        layout.addWidget(view_terms_button)
        layout.addSpacing(20)
        layout.addLayout(button_layout)

        dialog.setLayout(layout)
        result = dialog.exec()
        return result == QDialog.DialogCode.Accepted

    @pyqtSlot(str, str)
    def go_to_cell(self, sheet, address):
        assert(self.workbook is not None)
        self.workbook.activate(steal_focus=True)
        self.workbook.sheets[sheet].select()
        self.workbook.sheets[sheet].range(address).select()

    @pyqtSlot(str, list, str, tuple)
    def update_students(self, file_path, new_students, school_name, generated_date):
        self.pdf_path = file_path
        self.students = new_students
        self.school_name = school_name
        self.date_select.setDate(QDate(generated_date[2], generated_date[0], generated_date[1]))
        self.check_files_ready(did_update=True)

    @pyqtSlot(str)
    def update_workbook(self, new_workbook):
        self.workbook = xlwings.Book(new_workbook)
        # Update sheets in combo box

        self.update_sheet_selector()
        self.check_files_ready(did_update=True)

    @pyqtSlot(str, object)
    def update_docx(self, file_path, template):
        self.docx_path = file_path
        self.docx_template = template
        self.check_files_ready(did_update=True)

    def check_files_ready(self, did_update=False):
        has_students = bool(self.students)
        has_docx = bool(self.docx_path)

        # Check if excel window currently exists; clear if the window has been closed
        if self.workbook and self.workbook.fullname not in [i.fullname for i in xw.books]:
           self.workbook = None

        has_workbook = bool(self.workbook)

        # Grey out the add report to sheet button unless all data has been loaded
        self.add_absences_button.setEnabled(has_students and has_workbook)
        # Grey out the open pdf in new window button unless a pdf has been selected
        self.open_pdf_button.setEnabled(has_students)
        # Grey out the generate letter button unless doc and excel selected
        self.generate_template_button.setEnabled(has_workbook and has_docx)

        # Set checkboxes for each step
        self.step_containers[0].setTitle("1. " + ("☑" if has_workbook else "☐"))
        self.step_containers[1].setTitle("2. " + ("☑" if has_students else "☐"))
        self.step_containers[4].setTitle("4. " + ("☑" if has_docx else "☐"))

        # Set dropdown to sheet that best matches school name
        if did_update:
            self.step_containers[2].setTitle("3. ☐ (requires steps 1 and 2)")
            if has_workbook and has_students:
                best_sheet = self.best_match(self.school_name, [x.name for x in self.workbook.sheets])
                self.sheets_combo.setCurrentIndex(best_sheet + 1)

        return has_students and has_workbook

    def best_match(self, name, options):
        # Returns best matching string in options, -1 if none match well
        ratios = [SequenceMatcher(None, name.lower(), opt.lower()).ratio() for opt in options]
        maxr = max(ratios)
        if maxr > 0.5:
            return ratios.index(maxr)
        return -1

    def update_sheet_selector(self):
        self.sheets_combo.clear()
        if bool(self.workbook):
            self.sheets_combo.addItems(["[Create new]"] + [x.name for x in self.workbook.sheets])
