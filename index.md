---
title: Welcome to my blog
---

# Importing Libraries
import random
import sys
import openpyxl
from fpdf import FPDF
import os
import shutil
from copy import copy

# Customizing default values
random.seed(2022)

# Set the maximum depth of the Python interpreter stack to the required limit
sys.setrecursionlimit(2 ** 31 - 1)

# Give the location of the file
path = r"C:\Users\ritik\Desktop\current\timetable project\FinalInput.xlsx"
parent_dir = r"C:\Users\ritik\Desktop\current\timetable project"
directory1 = "Faculty_Wise"
directory2 = "Yearwise_Wise"
faculty_excel_path = os.path.join(parent_dir, directory1)
year_excel_path = os.path.join(parent_dir, directory2)

# For removing the Directories build from previous run session
try:
    shutil.rmtree(year_excel_path)
except:
    pass
try:
    shutil.rmtree(faculty_excel_path)
except:
    pass
# For creating the new Directories
# Here we will store yearwise timetable (which can be provided to the respective classes and students)
os.mkdir(year_excel_path)
#  Here we will store facultywise timetable (which can be provided to the individual faculties for their use)
os.mkdir(faculty_excel_path)


# Now we are going to read the input from the provided Excel file using openpyxl library

# Workbook object is created
wb_obj = openpyxl.load_workbook(path)

# Get workbook active sheet object
sheet_obj = wb_obj.active

# Get the total number of rows and columns present in the file
rows = sheet_obj.max_row
cols = sheet_obj.max_column

# Now we are going to store the entire file data into a dictionary named data
data = {i: [] for i in range(1, cols + 1)}
for i in range(2, rows + 1):
    for j in range(1, cols + 1):
        val = sheet_obj.cell(row=i, column=j).value
        try:
            item = int(val)
        except:
            item = val.strip()
        finally:
            data[j].append(item)

# print(data)

# Now we are storing the different information in seperate lists for future use
coursesCodes, facultyNames, courseyear, lecture, tutorial, practical = list(
    data.values()
)


# Defining Constants

# This data vary as per University/Organization needs which includes:
# WorkingDays : All the days in which classes can be scheduled
# Timings : Indicating all the possible timings in which classes can be scheduled
# Lunch or Break timing may also vary
WorkingDays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
Timings = [
    "9:00AM",
    "10:00AM",
    "11:00AM",
    "12:00PM",
    "1:00PM",
    "2:00PM",
    "3:00PM",
    "4:00PM",
    "5:00PM",
]
lunch = "1:00PM"
Ltime = Timings.index(lunch)

# print(i)

# Class timings before and after lunch break
beforelunch = Timings[:Ltime]
afterlunch = Timings[Ltime + 1 :]

# This file will be helpful in tracking the success or failure. It will also indicate the particular courses where failure occurs. Mainly serve the purpose for logs.
logs_generated = open("logs.txt", "w")


# There are three classes:
# Course class: corresponding to each course which is taught in the University


class course:
    # This is constructor
    def __init__(self, code, faculty=None, year=None, credits=(0, 0, 0)):
        self.courseCode = code
        self.faculty = faculty
        # Note that year is not an integer, its an object
        self.year = year
        self.L, self.T, self.P = credits

    # For later use (if any faculty wants to teach or wants to remain free on a particular slot)
    def specialallotment(self, day, time):
        pass


# Faculty class: corresponding to each faculty which is teaching in the University


class facultywiseSchedule:
    def __init__(self, name):
        """This is constructor for facultywiseSchedule class"""
        self.name = name
        self.classes = []
        self.schedule = {}
        for day in WorkingDays:
            self.schedule[day] = {}
            for time in Timings:
                self.schedule[day][time] = 0

    def addcourse(self, c):
        """Add classes/courses which the faculty will taught"""
        self.classes.append(c)

    def courses(self):
        """Returns the list of all the courses taught by this faculty"""
        return self.classes

    def check(self, CC, yearobj, day):
        """Checks whether faculty has taught CC course on this day"""
        for c in self.schedule[day]:
            if self.schedule[day][c] == CC:
                return True
        return False

    def addPClass(self, CC, yearobj, day):
        """Adds Practical Class (which is of 2 hrs) both in the faculty schedule and yearwise timetable"""
        timings = afterlunch
        i = WorkingDays.index(day)
        for j in range(
            2
        ):  #                     2 times since it checks for 2-4 then 3-5
            if not self.check(CC, yearobj, day):
                time = timings[j]
                ntime = timings[j + 1]
                if (
                    self.isavailable(day, time)
                    and self.isavailable(day, ntime)
                    and yearobj.slotavailable(day, time)
                    and yearobj.slotavailable(day, ntime)
                ):
                    self.setstatus(day, time, CC)
                    self.setstatus(day, ntime, CC)
                    yearobj.setstatus(day, time, self.name, CC)
                    yearobj.setstatus(day, ntime, self.name, CC)
                    print(
                        f"{CC} course Practical by faculty {self.name} - assigned succesfully"
                    )
                    logs_generated.write(
                        f"{CC} course Practical by faculty {self.name} - assigned succesfully\n"
                    )
                    return i
                else:
                    i = WorkingDays.index(day)
                    return self.addPClass(
                        CC, yearobj, WorkingDays[(i + 1) % len(WorkingDays)]
                    )
            else:
                return self.addPClass(
                    CC, yearobj, WorkingDays[(i + 1) % len(WorkingDays)]
                )

    def addClass(self, CC, yearobj, day, ltype, extraTime=False):
        """Adds Lecture/Tutorial Class (which is of 1 hrs) both in the faculty schedule and yearwise timetable"""
        flag = False
        if True:
            if extraTime == False:
                timings = beforelunch[:]
            else:
                timings = beforelunch[:] + ["2:00PM", "3:00PM", "4:00PM"]
            
            while timings != [] and not flag:
                # time = timings[0]
                time = random.choice(
                    timings
                )  # tries to assign to a random time slot of the day
                if yearobj.slotavailable(day, time) and self.isavailable(day, time):
                    self.setstatus(day, time, CC, ltype)
                    yearobj.setstatus(day, time, self.name, CC, ltype)
                    lec = "L" if ltype else "T"
                    print(
                        f"{CC} course {lec} by faculty {self.name} - assigned succesfully"
                    )
                    logs_generated.write(
                        f"{CC} course {lec} by faculty {self.name} - assigned succesfully\n"
                    )
                    flag = True  # when slot assigned successfully
                    break
                else:
                    timings.remove(
                        time
                    )  # since not possible in this time slot thus, next time check in other time slots only.

        return flag

    def setPstatus(self, day, time, ntime, mesg):
        """Book the slot in the faculty timetable for the Practical Class"""
        self.setstatus(day, time, mesg)
        self.setstatus(day, ntime, mesg)

    def setstatus(self, day, time, mesg, ltype="Practical"):
        """Book the slot in the faculty timetable for the Lecture/Tutorial Class"""
        lec = "Lecture" if ltype == 1 else "Tutorial" if ltype == 0 else ltype
        self.schedule[day][time] = mesg + "-" + lec

    def addClass_helper(self, CC, yearobj, d, ltype):
        """A helper method adding Lecture/Tutorial Class"""
        originald = d
        for (
            day
        ) in (
            WorkingDays
        ):  # tries to assign course starting in order Mon->Tue->....->Fri
            flag = self.addClass(CC, yearobj, WorkingDays[d], ltype)
            if flag:
                return d
            d = (d + 1) % len(
                WorkingDays
            )  # d corresponds to the day on which the course was assigned successfully in year's and faculty's timetables.
        d = originald
        for day in WorkingDays:  # tries to assign starting from Mon->Tue->....->Fri
            flag = self.addClass(CC, yearobj, WorkingDays[d], ltype, extraTime=True)
            if flag:
                return d
            d = (d + 1) % len(WorkingDays)

        # Failed to assign the course now, try backtracking
        

        # Still Failed
        # Print on the Console/Terminal
        print(f"{CC} course L/T by faculty {self.name} - CAN'T be assigned succesfully")

        # Write to the txt file (useful for logging information)
        logs_generated.write(
            f"{CC} course L/T by faculty {self.name} - CAN'T be assigned succesfully\n"
        )
        return -1

    def isavailable(self, day, time):
        """Returns True if faculty is available at a particular time slot on a particular day False otherwise"""
        return self.schedule[day][time] == 0

    def assignbreak(self, breakTiming):
        """Adds Lunch timings in the faculty timetable"""
        for day in self.schedule:
            self.schedule[day][breakTiming] = "LUNCH"


class yearwiseSchedule:
    """Branch/Batch/Section class: Corresponding to each year and each Course taught in the University"""

    def __init__(self, credits=[]):
        """This is constructor for yearwiseSchedule class"""
        self.timetable = {}
        for day in WorkingDays:
            self.timetable[day] = {}
            for time in Timings:
                self.timetable[day][time] = 0
        self.L, self.T, self.P = 0, 0, 0
        for L, T, P in credits:
            self.L = L
            self.T = T
            self.P = P

    def lunchtiming(self, lunch):
        """Adds Lunch timings in the Class timetable"""
        for d in WorkingDays:
            self.timetable[d][lunch] = "LUNCH"

    def slotavailable(self, day, time):
        """Returns True if Class timetable is free at a particular time slot on a particular day False otherwise"""
        return self.timetable[day][time] == 0

    def setstatus(self, day, time, fac, cc, ltype="Practical"):
        """Book the slot in the Class timetable"""
        lec = "Lecture" if ltype == 1 else "Tutorial" if ltype == 0 else ltype
        self.timetable[day][time] = (fac, cc + "-" + lec)


years = 3
# For the time being we are focussing on the Bsc Mathematics in order to avoid the complexity (although this can be extended for any branch and any stream)
text = "BSc Maths "

# bscmath is a dict with keys being integers (e.g. 1,2,3) representing the BSc Math Year and values being the objects of class yearwiseSchedule corresponding to these years
bscmath = {}
for i in range(1, years + 1):
    y = yearwiseSchedule()
    y.lunchtiming(lunch)
    bscmath[i] = y

# FN is a list containing unique faculty names
FN = list(set(facultyNames))
f = len(FN)
# faculties is a dict with keys being Strings (e.g. Samuel) representing the Name of the faculty and values being the objects of class facultywiseSchedule corresponding to the faculty
faculties = {}
for i in range(f):
    newfaculty = facultywiseSchedule(FN[i])
    faculties[FN[i]] = newfaculty

n = len(coursesCodes)
# courseobj is a dict with keys being Strings (e.g. MAT217) representing the Course Code and values being the objects of class course corresponding to that course code
courseobj = {}
for i in range(n):
    newcourse = course(  # here the courses
        coursesCodes[i],
        facultyNames[i],
        bscmath[courseyear[i]],
        credits=(lecture[i], tutorial[i], practical[i]),
    )
    courseobj[coursesCodes[i]] = newcourse
    faculties[facultyNames[i]].addcourse(newcourse)


# __________________________________________________________________________________________________________________
# THE HEART of the entire program lies here

# It runs for all the courses (i.e. Subjects) and allocate it in the slot which is available both in the Faculty timetable and the Yearwise timetable

for times in range(2):
    for C in courseobj.values():
        d = 0
        L, T, P = C.L, C.T, C.P
        F = faculties[C.faculty]
        if times == 0:
            for i in range(P):
                D = WorkingDays[d % len(WorkingDays)]
                d = F.addPClass(C.courseCode, C.year, D)
                d += 1
        if times == 1:
            d = 0
            for i in range(L + T):
                D = WorkingDays[d % len(WorkingDays)]
                d = F.addClass_helper(C.courseCode, C.year, d, ltype=i < L) + 1
                d = d % len(WorkingDays)
# __________________________________________________________________________________________________________________


# Now closing the logs file.
logs_generated.close()

# This entire block is for Generating Aesthetic Output on the terminal
# Most of the statements and numbers inside f-strings are just for aesthetic purpose and can be ignored
for i in bscmath.keys():
    print("\n")
    print("_" * 180)
    print("\n")
    print(text + str(i) + " Year Timetable")
    d = bscmath[i].timetable
    t = "Days/Timings"
    print(f"{t:15s}", end="")
    for i in range(len(Timings) - 1):
        s = Timings[i] + "- " + Timings[i + 1]
        print(f"{s:25s}", end="")
    print("\n")
    for day in WorkingDays:
        print("\n")
        print(f"{day:15s}", end="")
        l = list(d[day].values())
        for k in range(len(l) - 1):
            v = l[k]
            if v == "LUNCH":
                pass
            elif v == 0:
                v = "---"
            else:
                f, c = v
                if len(f) > 10:
                    f = f[:10] + ".."
                t = 7
                if c[8] in "LTP":
                    t = 8
                v = f[:13] + "[" + c[:t] + "]"
            print(f"{v:25s}", end="")
    print()


def render_table_header(l=["Faculty", "Time"]):
    """Used for printing the header for the PDF files"""
    pdf.set_font(style="B")  # enabling bold text
    for col_name in TABLE_COL_NAMES:
        if col_name in l:
            pdf.cell(col_width * 1.5, line_height, col_name, border=1)
        else:
            pdf.cell(col_width, line_height, col_name, border=1)
    pdf.ln(line_height)
    pdf.set_font(style="")  # disabling bold text


def copy_sheet(source_sheet, target_sheet):
    copy_cells(source_sheet, target_sheet)  # copy all the cel values and styles
    copy_sheet_attributes(source_sheet, target_sheet)


def copy_sheet_attributes(source_sheet, target_sheet):
    target_sheet.sheet_format = copy(source_sheet.sheet_format)
    target_sheet.sheet_properties = copy(source_sheet.sheet_properties)
    target_sheet.merged_cells = copy(source_sheet.merged_cells)
    target_sheet.page_margins = copy(source_sheet.page_margins)
    target_sheet.freeze_panes = copy(source_sheet.freeze_panes)

    # set row dimensions
    # So you cannot copy the row_dimensions attribute. Does not work (because of meta data in the attribute I think). So we copy every row's row_dimensions. That seems to work.
    for rn in range(len(source_sheet.row_dimensions)):
        target_sheet.row_dimensions[rn] = copy(source_sheet.row_dimensions[rn])

    if source_sheet.sheet_format.defaultColWidth is None:
        # print('Unable to copy default column wide')
        pass
    else:
        target_sheet.sheet_format.defaultColWidth = copy(
            source_sheet.sheet_format.defaultColWidth
        )

    # set specific column width and hidden property
    # we cannot copy the entire column_dimensions attribute so we copy selected attributes
    for key, value in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[key].min = copy(
            source_sheet.column_dimensions[key].min
        )  # Excel actually groups multiple columns under 1 key. Use the min max attribute to also group the columns in the targetSheet
        target_sheet.column_dimensions[key].max = copy(
            source_sheet.column_dimensions[key].max
        )  # https://stackoverflow.com/questions/36417278/openpyxl-can-not-read-consecutive-hidden-columns discussed the issue. Note that this is also the case for the width, not only the hidden property
        target_sheet.column_dimensions[key].width = copy(
            source_sheet.column_dimensions[key].width
        )  # set width for every column
        target_sheet.column_dimensions[key].hidden = copy(
            source_sheet.column_dimensions[key].hidden
        )


def copy_cells(source_sheet, target_sheet):
    for (row, col), source_cell in source_sheet._cells.items():
        target_cell = target_sheet.cell(column=col, row=row)

        target_cell._value = source_cell._value
        target_cell.data_type = source_cell.data_type

        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)

        if source_cell.hyperlink:
            target_cell._hyperlink = copy(source_cell.hyperlink)

        if source_cell.comment:
            target_cell.comment = copy(source_cell.comment)


# Creating a new Excel Workbook
wb = openpyxl.Workbook()
sheet = wb.active

# This will be the header for our PDFs and Excel files
TABLE_COL_NAMES = ["Course Code", "Faculty", "Day", "Time", "Lecture Type"]

# Utility Dictionary - Would help in getting the correct suffix part in the file names
fnameutility = {1: "st", 2: "nd", 3: "rd"}

# Generating Output Year-wise
for i in range(1, 4):
    fname = str(i) + fnameutility[i]
    sheet.title = fname + " Year"
    for j in range(5):
        sheet.cell(1, j + 1).value = TABLE_COL_NAMES[j]
    Y = bscmath[i]
    row = 2
    Timings.append("5:00PM")
    TABLE_DATA = []
    for day in WorkingDays:
        for time in Timings:
            if time != Timings[-1]:
                ntime = Timings[Timings.index(time) + 1]
                if Y.timetable[day][time] != "LUNCH" and Y.timetable[day][time] != 0:
                    F, mesg = Y.timetable[day][time]
                    try:
                        CC, ltype = mesg.split("-")
                    except:
                        continue
                    sheet.cell(row, 1).value = CC
                    sheet.cell(row, 2).value = F
                    sheet.cell(row, 3).value = day
                    sheet.cell(row, 4).value = f"{time} - {ntime}"
                    if time in beforelunch:
                        Ltype = ltype
                    else:
                        Ltype = "Practical"
                    sheet.cell(row, 5).value = Ltype
                    TABLE_DATA.append(
                        [CC, F, day, f"{time} - {ntime}", sheet.cell(row, 5).value]
                    )
                    row += 1
    if i != 3:
        sheet = wb.create_sheet()
    # Generate Excel file with the correct name
    wb_target = openpyxl.Workbook()
    target_sheet = wb_target.active
    target_sheet.title = fname
    source_sheet = sheet
    copy_sheet(source_sheet, target_sheet)
    wb_target.save(f"{year_excel_path}\BScMath_{fname}_Year_TimeTable.xlsx")

    # # Create a new PDF
    # pdf = FPDF()
    # pdf.add_page()
    # pdf.set_font("Times", size=12)
    # line_height = pdf.font_size * 2
    # col_width = pdf.epw / (len(TABLE_COL_NAMES) + 1)  # distribute content evenly

    # # Call the header funtion
    # render_table_header()
    # for row in TABLE_DATA:
    #     if pdf.will_page_break(line_height):
    #         render_table_header()
    #     for datum in row:
    #         if row[1] == datum or row[3] == datum:
    #             pdf.cell(col_width * 1.5, line_height, datum, border=1)
    #         else:
    #             pdf.cell(col_width, line_height, datum, border=1)
    #     pdf.ln(line_height)

    # # Generate PDF with the correct file name
    # pdf.output(f"{year_excel_path}\TimeTable_BScMath_{fname}_Year.pdf")
# Generate Compiled Excel file


wb.save(f"{year_excel_path}\BScMath_Compiled_TimeTable.xlsx")


# Again creating a new Excel Workbook
wb = openpyxl.Workbook()
sheet = wb.active

# This will be the header for our PDFs and Excel files
TABLE_COL_NAMES = ["Course Code", "Year", "Day", "Time", "Lecture Type"]

# Generating Output Faculty-wise
for i in facultyNames:
    fname = faculties[i].name
    sheet.title = fname
    for j in range(5):
        sheet.cell(1, j + 1).value = TABLE_COL_NAMES[j]
    F = faculties[i]
    row = 2
    TABLE_DATA = []
    for day in WorkingDays:
        for time in Timings:
            if time != Timings[-1]:
                ntime = Timings[Timings.index(time) + 1]
                if F.schedule[day][time] != "LUNCH" and F.schedule[day][time] != 0:
                    mesg = F.schedule[day][time]
                    try:
                        CC, ltype = mesg.split("-")
                    except:
                        continue
                    sheet.cell(row, 1).value = CC
                    sheet.cell(row, 2).value = courseyear[coursesCodes.index(CC)]
                    sheet.cell(row, 3).value = day
                    sheet.cell(row, 4).value = f"{time} - {ntime}"
                    if time in beforelunch:
                        Ltype = ltype
                    else:
                        Ltype = "Practical"
                    sheet.cell(row, 5).value = Ltype
                    TABLE_DATA.append(
                        [
                            CC,
                            courseyear[coursesCodes.index(CC)],
                            day,
                            f"{time} - {ntime}",
                            sheet.cell(row, 5).value,
                        ]
                    )
                    row += 1

    # Generate Excel file with the correct name
    wb_target = openpyxl.Workbook()
    target_sheet = wb_target.active
    target_sheet.title = fname
    source_sheet = sheet
    copy_sheet(source_sheet, target_sheet)

    wb_target.save(f"{faculty_excel_path}\{fname}_TimeTable.xlsx")
    sheet = wb.create_sheet()

    # pdf = FPDF()
    # pdf.add_page()
    # pdf.set_font("Times", size=12)
    # line_height = pdf.font_size * 2
    # col_width = pdf.epw / (len(TABLE_COL_NAMES) + 1)  # distribute content evenly

    # render_table_header()
    # for row in TABLE_DATA:
    #     if pdf.will_page_break(line_height):
    #         render_table_header()
    #     for datum in row:
    #         if row[3] == datum:
    #             pdf.cell(col_width * 1.5, line_height, datum, border=1)
    #         else:
    #             pdf.cell(col_width, line_height, str(datum), border=1)
    #     pdf.ln(line_height)

    # # Generate PDF with the correct file name
    # pdf.output(f"{faculty_excel_path}\TimeTable_{fname}_.pdf")

# Generate Compiled Excel file
wb.save(f"{faculty_excel_path}\Faculty_Compiled_TimeTable.xlsx")
