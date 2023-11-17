
import csv
import openpyxl


class Section:
    def __init__(self, section_type, days, time,room):
        #Initialising
        self.section_type = section_type
        self.days = days.split(',')
        self.time = time
        self.room = room
        
    def get_schedule(self):
        return f" Days: {self.days}, Hour : {self.time}"
    
class LectureSection(Section):
    def __init__(self, name, schedule, room):
        super().__init__(name, schedule)
        self.room = room

class LabSection(Section):
    def __init__(self, name, schedule, lab_number):
        super().__init__(name, schedule)
        self.lab_number = lab_number

class TutorialSection(Section):
    def __init__(self, name, schedule, tutor):
        super().__init__(name, schedule)
        self.tutor = tutor
class Course:
    __admin_pass = None

    def __init__(self,password ,course_code, course_name, exam_date=[]):
        self.course_code = course_code
        self.course_name = course_name
        self.exam_date = exam_date
        self.__admin_pass = password
        self.sections = {}

    def get_sections(self):
        return list(self.sections.keys())

    def __str__(self):
        #Prints basic info about the Course object
        print(f"Course Code: {self.course_code}")
        print(f"Course Name: {self.course_name}")
        print(f"Exam Dates: {', '.join(self.exam_date)}")
        print("Sections:")
        #Iterates through all the subjects
        for section, details in self.sections.items():
            print(f"  {section}: {details['instructor']} - {details['timing']}")
            print(f"     Schedule: {details['section'].get_schedule()}")

    #A private(__ dunder used) method to add a section
    def __add_section(self, section, instructor, section_type, days, time, room):
        #Checking if the section to be added already exists
        if section not in self.sections:
            self.sections[section] = {'instructor': instructor, 'details': Section(section_type, days, time, room)}
            print(f"Section {section} added successfully.")
        else:
            print(f"Section {section} already exists. Choose a different section.")

    
    def populate_sections(self):
        # A simple password authentication to access the __add_secton private method
        
        password = int(input("Enter the admin password to add new sections: "))
        if password != self.__admin_pass:
            print("Access denied. Incorrect admin password.")
            return

        section_type = input("Enter the section type (e.g., L, P, T): ")
        time = []
        if section_type == "L":
            time.append(int(input("Enter slot(hour) of the section: ")))
            
        elif section_type == "P":
            x = int(input("Enter starting slot of the section: "))
            time.append(x)
            time.append(x+1)
            if self.course_code == "MEF112":
                time.append(x+2)
        elif section_type == "T":
            time.append(int(input("Enter slot(hour) of the section: ")))
        else:
            print("!!!Please enter a valid section type!!!")   
            return 
        section = input("Enter section code:")
        instructor = input("Enter Instructor name:")
        days = input("Enter the days(s) of the week: (MO,TU,WD,TH,FR,SA comma separated for multiple)")
        
        
        room = int(input("Enter room no: "))

        self.__add_section( section, instructor, section_type, days, time,room)


class Timetable:
    def __init__(self):
        self.subjects = {}
        self.table = {'MO': ['']*9, 'TU': ['']*9, 'WD': ['']*9, 'TH': ['']*9, 'FR': ['']*9, 'SA': ['']*9}
    def enroll_subject(self,course):
        if course.course_code not in self.subjects:
            self.subjects[course.course_code] = {'course': course, 'sections' : []}
            print(f"Enrolled in {course.course_code} successfully.")
        else:
           print(f"You are already enrolled in {course.course_code}")

    def add_section_to_table(self):
        for course_code,course in self.subjects.items():
            for section_code,section in course['course'].sections.items():
                for day in section['details'].days:
                    for time in section['details'].time:
                        if time not in self.table[day]:
                            self.table[day][time] = f"{section_code} {course_code}"
                        else:
                            print(f"Section :{section_code} Course : {course_code} is already in {day} in hour {time}")

    def print_timetable(self):
        print("   | 8:00 AM | 9:00 AM | 10:00 AM | 11:00 AM | 12:00 PM | 1:00 PM | 2:00 PM | 3:00 PM | 4:00 PM |")
        print("--------------------------------------------------")
        for day in ['MO', 'TU', 'WD', 'TH', 'FR', 'SA']:
            print(f"{day} | ", end="")
            for time_slot in self.table[day]:
                print(f"{time_slot:<9} | ", end="")
            print()

            
    def check_clashes(self):
        # Check for clashes in examination dates
        exam_dates = set()
        for subject_data in self.subjects.values():
            exam_dates.add(subject_data['course'].exam_date)

        if len(exam_dates) != len(self.subjects):
            print("Clash found in examination dates.")
        else:
            print("No clash in examination dates.")

        # Check for clashes in section timings
        section_schedule = {}
        for subject_data in self.subjects.values():
            for section_data in subject_data['sections']:
                section = section_data['section']
                if section.get_schedule() in section_schedule:
                    print(f"Clash found in section schedule: {section.get_schedule()}")
                else:
                    section_schedule[section.get_schedule()] = subject_data['course'].course_code
            

# Example Usage:

    def export_to_csv(self, filename='timetable.csv'):
        with open(filename, 'w', newline='') as csvfile:
            fieldnames = ['Time'] + list(self.table.keys())
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            writer.writeheader()
            for time_slot in range(0, 9):  
                row_data = {'Time': f'{time_slot}:00'}
                for day in self.table.keys():
                    row_data[day] = self.table[day][time_slot] if self.table[day][time_slot] else ' '
                writer.writerow(row_data)


def populate_subject(excel_file_path, timetable,password):
    COLUMN_COURSE_CODE = 2  # Assuming 'Code' is the first column
    COLUMN_COURSE_NAME = 3  # Assuming 'Name' is the second column
    COLUMN_EXAM_DATES = 4  # Assuming 'Exam Date' is the third column

    workbook = None

    try:
        workbook = openpyxl.load_workbook(excel_file_path)
        sheet = workbook.active  # You can replace this with specific sheet access if needed

        for row in sheet.iter_rows(min_row=2, values_only=True):
            course_code = row[COLUMN_COURSE_CODE - 1]
            course_name = row[COLUMN_COURSE_NAME - 1]
            exam_dates = row[COLUMN_EXAM_DATES - 1]

            # Create a new Course instance and add it to the timetable
            new_course = Course(password,course_code, course_name, exam_dates)
            timetable.enroll_subject(new_course)

        print("Subjects populated successfully.")

    except Exception as e:
        print(f"Error while populating subjects: {e}")

    finally:
        if workbook:
            workbook.close()

# Example Usage:
timetable = Timetable()

# Adjust the file path based on your scenario
excel_file_path = "/home/yahboi0lem/CSstuff/Backend/DVM/T1_ttmanage/Book1.xlsx"

password = int(input(print("Set your pin:")))
populate_subject(excel_file_path, timetable,password)

timetable.check_clashes()

flag = True

while(flag):
    choice = int(input(print("Enter choice: \n 1:Enter a new section \n 2:Save your changes\n 3:Show Time Table \n 4:Export Time Table \n 0:QUIT")))
    if choice == 1:
        print("Enter Course code to add a section in: \n")
        for course in timetable.subjects.keys():
            print(f"{str(course)} ")
        course_str = input("->")
        if course_str not in timetable.subjects.keys():
            print("!!!Please enter a valid code!!!")
            continue
        else:
            course = timetable.subjects[course_str]['course']
            course.populate_sections()
    elif choice ==  2:
        print("Adding your sections to your timetable......")
        timetable.add_section_to_table()
    elif choice == 4:    
        try :
            timetable.export_to_csv()
            print("Timetable exported successfully")
        except Exception as e:
            print(f"Time table unable to load error: {e}")

    elif choice == 3:
        timetable.print_timetable()        
    elif choice == 0:
        flag = False
        timetable.add_section_to_table()
        print("")
        print("Quitting.....")
        
        break 
    else:
        print("Please enter a valid option")
        continue  



