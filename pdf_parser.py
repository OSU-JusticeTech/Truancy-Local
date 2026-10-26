import logging
import re
from constructor import Student
import pdfplumber

logger = logging.getLogger("pdf_parser")

logger.setLevel(logging.DEBUG)

def _student_summary(student):
    return (f"id={getattr(student, 'id', None)!r} "
            f"name={getattr(student, 'lastName', None)!r}, {getattr(student, 'firstName', None)!r}")

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

    logger.info("Starting parse of %r", pdf_path)

    #open and process each page in the PDF
    with pdfplumber.open(pdf_path) as pdf:
        logger.info("PDF opened: %d page(s)", len(pdf.pages))
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                #skip pages with no content
                logger.warning("Page %d: no text extracted (blank or scanned image?), skipping", page_num)
                continue
            lines = text.split("\n")
            logger.debug("Page %d: %d line(s) extracted", page_num, len(lines))
            for line_num, line in enumerate(lines, start=1):
                where = f"p{page_num}:l{line_num}"
                logger.debug("[%s] %rx", where, line)

                # set school name if this line contains it
                school_match = re.match(school_pattern, line)
                if school_match:
                    logger.debug("[%s] school line matched: %r", where, line)
                    assert not school_name or school_name == line
                    school_name = line

                # set date if this line contains one
                date_match = re.match(date_pattern, line)
                if date_match:
                    generated_date = (int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3)))
                    logger.debug("[%s] date matched: %s", where, generated_date)

                # start of a new student block based on the # symbol
                if "#" in line:
                    logger.debug("[%s] '#' found, starting new student block", where)
                    #if student was already being parsed, save it
                    if current_student:
                        logger.debug("Student %s complete", _student_summary(current_student))
                        students.append(current_student)
                    #new student object
                    current_student = Student.create_empty()
                    try:
                        parts = line.split("#")
                        name_part = parts[0].strip()
                        student_id_part = parts[1].strip()
                        #parse name/id from line
                        if "," in name_part:
                            name_split = name_part.split(",")
                            current_student.firstName = name_split[1].strip()
                            current_student.lastName = name_split[0].strip()
                        else:
                            name_split = name_part.split()
                            current_student.firstName = name_split[0]
                            current_student.lastName = " ".join(name_split[1:])
                        current_student.id = student_id_part.split()[0]
                        logger.debug("extracted student %s %s #%s", current_student.firstName,
                                     current_student.lastName, current_student.id)
                    except IndexError:
                        logger.error("[%s] could not parse student header line: %r "
                                     "(parts=%r)", where, line, line.split("#"))
                        raise

                # extract other fields while inside a student block
                if current_student:
                    #match age if line has
                    age_match = age_pattern.search(line)
                    if age_match:
                        current_student.age = age_match.group(1)
                        logger.debug("[%s] age=%s", where, current_student.age)
                    #match grade if on line
                    grade_match = grade_pattern.search(line)
                    if grade_match:
                        current_student.grade = grade_match.group(1)
                        logger.debug("[%s] grade=%s", where, current_student.grade)
                    #match attendence data values
                    attn_match = attendance_pattern.search(line)
                    if attn_match:
                        logger.debug("[%s] attendance row matched, raw groups=%r",
                                     where, attn_match.groups())
                        try:
                            current_student.excused = str(float(attn_match.group(1).strip()))
                            current_student.unexcused = str(float(attn_match.group(2).strip()))
                            current_student.medical = str(float(attn_match.group(3).strip()))
                            current_student.suspension = str(float(attn_match.group(4).strip()))
                            current_student.schoolTotal = str(float(attn_match.group(5).strip()))
                            current_student.attendingTotal = str(float(attn_match.group(6).strip()))
                            current_student.absenceTotal = str(float(attn_match.group(7).strip()))
                        except ValueError:
                            logger.error("[%s] non-numeric attendance value for %s; line=%r groups=%r",
                                         where, _student_summary(current_student), line,
                                         attn_match.groups())
                            raise
                elif age_pattern.search(line) or grade_pattern.search(line) or attendance_pattern.search(line):
                    logger.warning("[%s] student-data line found before any student header "
                                   "(ignored): %r", where, line)

        # add final student
        if current_student:
            logger.debug("Student %s complete", _student_summary(current_student))
            students.append(current_student)

    logger.info("Parse complete: school=%r date=%s students=%d",
                school_name, generated_date, len(students))

    return school_name, generated_date, students
