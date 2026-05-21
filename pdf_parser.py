import re
from constructor import Student
import pdfplumber

# function to parse students from a truancy report PDF
def extract_students_from_pdf(pdf_path):
    students = []
    current_student = None

    # regex patterns
    age_pattern = re.compile(r'Age:\s*(\d+)')
    grade_pattern = re.compile(r'Grade:\s*([A-Za-z0-9]+)')
    attendance_pattern = re.compile(
        r'\d\d\d\d-\d\d\d\d\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)'
    )
    school_pattern = re.compile(r'\D* \d\d\d\d-\d\d\d\d')
    date_pattern = re.compile(r'(\d\d)/(\d\d)/(\d\d\d\d)')

    school_name = ""
    generated_date = None

    #open and process each page in the PDF
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                #skip pages with no content
                continue
            lines = text.split("\n")
            for line in lines:
                # set school name if this line contains it
                school_match = re.match(school_pattern, line)
                if school_match:
                    assert not school_name or school_name == line
                    school_name = line

                # set date if this line contains one
                date_match = re.match(date_pattern, line)
                if date_match:
                    generated_date = (int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3)))

                # start of a new student block based on the # symbol
                if "#" in line:
                    #if student was already being parsed, save it
                    if current_student:
                        students.append(current_student)
                    #new student object
                    current_student = Student.create_empty()
                    parts = line.split("#")
                    name_part = parts[0].strip()
                    student_id_part = parts[1].strip()
                    #parse name/id from line
                    name_split = name_part.split(",")
                    current_student.firstName = name_split[1].strip()
                    current_student.lastName = name_split[0].strip()
                    current_student.id = student_id_part.split()[0]

                # extract other fields while inside a student block
                if current_student:
                    #match age if line has
                    age_match = age_pattern.search(line)
                    if age_match:
                        current_student.age = age_match.group(1)
                    #match grade if on line
                    grade_match = grade_pattern.search(line)
                    if grade_match:
                        current_student.grade = grade_match.group(1)
                    #match attendence data values
                    attn_match = attendance_pattern.search(line)
                    if attn_match:
                        current_student.excused = str(float(attn_match.group(1).strip()))
                        current_student.unexcused = str(float(attn_match.group(2).strip()))
                        current_student.medical = str(float(attn_match.group(3).strip()))
                        current_student.suspension = str(float(attn_match.group(4).strip()))
                        current_student.schoolTotal = str(float(attn_match.group(5).strip()))
                        current_student.attendingTotal = str(float(attn_match.group(6).strip()))
                        current_student.absenceTotal = str(float(attn_match.group(7).strip()))

        # add final student
        if current_student:
            students.append(current_student)

    return school_name, generated_date, students
