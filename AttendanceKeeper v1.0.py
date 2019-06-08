from Tkinter import *
import ttk, tkFileDialog, Tkconstants
import xlrd
from xlwt import Workbook


class Student:
    def __init__(self, id, name, department, section):
        self.id = id
        self.department = department
        self.section = section
        (self.first_name, self.surname) = name.split(' ', 1)  # splitting the name by first space only

    def __repr__(self):  # class representation
        return self.surname + ', ' + self.first_name + ', ' + str(self.id)


class StudentList():
    def __init__(self, list_name):
        self.list_name = list_name
        self.students = []

    def __repr__(self):  # class representation
        return self.list_name


class AttendanceTool(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent
        self.extentions = ['txt', 'xls']
        self.initUI()

    def initUI(self):  # UI Initialization

        self.title_label = Label(self, text='AttendanceKeeper v1.0', font='Helvetica 16 bold')
        self.student_list_label = Label(self, text='Select student list Excel file:', font='Helvetica 10 bold')
        self.import_button = Button(self, text='Import List', font='Helvetica 10 bold', command=self.import_function)
        self.select_student_label = Label(self, text='Select a Student:', font='Helvetica 10 bold')
        self.section_label = Label(self, text='Section:', font='Helvetica 10 bold')
        self.attendedstudents_label = Label(self, text='Attended Students:', font='Helvetica 10 bold')
        self.student_listbox = Listbox(self, selectmode='multiple', height=5, width=36)
        self.attended_listbox = Listbox(self, selectmode='multiple', height=5, width=36)
        self.section_combobox = ttk.Combobox(self, state='readonly')
        self.add_button = Button(self, text='Add =>', font='Helvetica 10 bold', command=self.add_student)
        self.remove_button = Button(self, text='Remove <=', font='Helvetica 10 bold', command=self.remove_student)
        self.filetype_label = Label(self, text='Please select file type:', font='Helvetica 10 bold')
        self.filetype_combobox = ttk.Combobox(self, state='readonly', width=5)
        self.weekentry_label = Label(self, text='Please enter week:', font='Helvetica 10 bold')
        self.weekentry = Entry(self)
        self.export_button = Button(self, text='Export as File', font='Helvetica 10 bold', command=self.export_file)

        self.scrollbar1 = Scrollbar(self)
        self.scrollbar2 = Scrollbar(self)
        self.student_listbox.config(yscrollcommand=self.scrollbar1.set)  # binding the scrollbars to the listboxes
        self.scrollbar1.config(command=self.student_listbox.yview)
        self.attended_listbox.config(yscrollcommand=self.scrollbar2.set)
        self.scrollbar2.config(command=self.attended_listbox.yview)

        self.filetype_combobox['values'] = self.extentions
        self.filetype_combobox.current(0)  # set default to txt (first value in list)

        self.pack()
        self.UIPacking()

    def UIPacking(self):  # Packing using grid with 4 rows 3 columns

        self.title_label.grid(in_=self, row=0, columnspan=3)

        self.student_list_label.grid(in_=self, row=1, column=0, sticky=W)
        self.import_button.grid(in_=self, row=1, column=1, sticky=W + E)

        self.select_student_label.grid(in_=self, row=2, column=0, sticky=W)
        self.section_label.grid(in_=self, row=2, column=1)
        self.attendedstudents_label.grid(in_=self, row=2, column=2, sticky=W)

        self.student_listbox.grid(in_=self, row=3, column=0, sticky=W + E)
        self.scrollbar1.grid(in_=self, row=3, column=0, sticky=E + N + S)
        self.section_combobox.grid(in_=self, row=3, column=1, stick=N)
        self.attended_listbox.grid(in_=self, row=3, column=2, sticky=W + E)
        self.scrollbar2.grid(in_=self, row=3, column=2, sticky=E + N + S)
        self.add_button.grid(in_=self, row=3, column=1, sticky=E + W)
        self.remove_button.grid(in_=self, row=3, column=1, sticky=E + W + S)

        self.filetype_label.grid(in_=self, row=4, column=0, sticky=W)
        self.filetype_combobox.grid(in_=self, row=4, column=0, sticky=E)
        self.weekentry_label.grid(in_=self, row=4, column=1, sticky=E)
        self.weekentry.grid(in_=self, row=4, column=2, sticky=W)
        self.export_button.grid(in_=self, row=4, column=2, sticky=E)

    def import_function(self):  # File importing and class objects creation

        imported_file = tkFileDialog.askopenfilename(title='Select student list',
                                                     filetypes=(("excel files", "*.xls*"), ("all files", "*.*")))
        workbook = xlrd.open_workbook(imported_file)
        sheet = workbook.sheet_by_index(0)

        getting_sections = {}  # temporary usage
        self.all_students = []  # this will contain all student objects

        for row in range(1, sheet.nrows):  # start reading from 2nd row
            getting_sections[
                str(sheet.cell_value(row, 3))] = []  # used to get sections (dictionary so that its not duplicated)

            # student object creation for all students in file
            self.all_students.append(
                Student(str(int(sheet.cell_value(row, 0))), sheet.cell_value(row, 1).encode('utf-8'),
                        str(sheet.cell_value(row, 2)), str(sheet.cell_value(row, 3))))
            # used int to get rid of the decimals but the actual value will be used a string later on in exporting

        self.all_students.sort(key=lambda x: x.surname)  # sorting the all students list using the attribute surname

        sections = []  # add the sections to a list now
        for key in getting_sections:
            sections.append(key)

        self.stored_sections = []  # store as a list of Studentlist objects
        for i in sections:
            self.stored_sections.append(StudentList(i))

        self.stored_sections.sort(key=lambda x: x.list_name)  # sort the sections by number

        for list in self.stored_sections:
            for student in self.all_students:
                if str(student.section) == str(
                        list.list_name):  # if the student's section matches the list , add him/her to it
                    list.students.append(student)
        list_for_adding = []  # to add sectoin names to the combobox , using representation in class definition
        for i in self.stored_sections:
            list_for_adding += i.__repr__(),
        self.section_combobox['values'] = list_for_adding
        self.section_combobox.current(0)  # set current value to first section

        # FIRST INSERTION ONLY
        first_value = self.stored_sections[0]
        for std in first_value.students:
            self.student_listbox.insert(END, std.__repr__(), )

        # event binding
        self.section_combobox.bind("<<ComboboxSelected>>", self.section_selection)

    def section_selection(self, event):
        self.student_listbox.delete(0, END)  # clear both upon reselection
        self.attended_listbox.delete(0, END)
        self.value_of_combo = self.section_combobox.get()
        for list in self.stored_sections:
            for std in list.students:
                if self.value_of_combo == list.list_name:
                    self.student_listbox.insert(END, std.__repr__(), )
                    # insert the students of a studentlist only if the combobox selection matches to the lists name

    def add_student(self):
        selected_index = map(int, self.student_listbox.curselection())  # returns a list of selected indices

        for i in selected_index:
            value = self.student_listbox.get(i)
            self.attended_listbox.insert(END, value)  # add to attended listbox
            self.student_listbox.delete(i)  # delete from students listbox
            for j in range(0, len(
                    selected_index)):  # index update for each loop (because were removing from one list so the index changes everytime by 1)
                selected_index[j] -= 1

    def remove_student(self):
        selected_index = map(int, self.attended_listbox.curselection())  # list of selected indices
        # just the opposite of before
        for i in selected_index:
            value = self.attended_listbox.get(i)
            self.student_listbox.insert(END, value)
            self.attended_listbox.delete(i)
            for j in range(0, len(selected_index)):  # index update for each loop
                selected_index[j] -= 1

    def export_file(self):
        filetype_selection = self.filetype_combobox.get()
        entry_value = self.weekentry.get()  # entry widget value

        if filetype_selection == 'txt':
            section_name = self.section_combobox.get()
            file_name = section_name + ' ' + entry_value + '.txt'
            with open(file_name, 'w+') as output_file:
                output_file.write('Id	Name	Dept.\n')  # first line which is static
                for i in self.attended_listbox.get(0, END):  # accessing the values of the listbox
                    (surname, name, id) = i.encode('utf-8').split(', ')  # spliting and encoding for turkish letters
                    output_file.write(id + ' ' + name + ' ' + surname + ' ')
                    for student in self.all_students:  # loop to get the department of the student
                        if id == student.id:
                            dept = student.department
                            output_file.write(dept + '\n')

        if filetype_selection == 'xls':
            section_name = self.section_combobox.get()
            file_name = section_name + ' ' + entry_value + '.xls'
            book = Workbook(encoding='utf-8')
            sheet = book.add_sheet('Attendance')

            # static lines
            sheet.row(0).write(0, 'Id')
            sheet.row(0).write(1, 'Name')
            sheet.row(0).write(2, 'Dept.')

            row_counter = 1  # updates rows

            for i in self.attended_listbox.get(0, END):
                (surname, name, id) = i.encode('utf-8').split(', ')
                sheet.row(row_counter).write(0, id)
                sheet.row(row_counter).write(1, name + ' ' + surname)
                for student in self.all_students:
                    if id == student.id:
                        dept = student.department
                        sheet.row(row_counter).write(2, dept)
                row_counter += 1

            book.save(file_name)


def main():
    root = Tk()
    root.title('Attendance Tool')
    app = AttendanceTool(root)
    root.mainloop()


if __name__ == '__main__':
    main()