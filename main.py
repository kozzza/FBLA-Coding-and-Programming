import sys
import sqlite3
import random
from faker import Faker
from datetime import datetime, timedelta, date
from fpdf import FPDF

from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QTableWidget, \
    QTableWidgetItem, QComboBox, QVBoxLayout, QAbstractScrollArea, QMessageBox, \
    QLineEdit, QLabel, QSpacerItem, QSizePolicy, QHeaderView, QTabWidget, QAbstractItemView, \
    QCalendarWidget, QFrame, QApplication, QGraphicsTextItem, QGraphicsScene, QGraphicsView, QGraphicsRectItem
from PyQt5.QtGui import QIcon, QColor
from PyQt5.QtCore import pyqtSlot, Qt, QDate, QRect, QPoint, QSize


#create table and instantiate cursor
conn = sqlite3.connect('communityservice.db')
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS studentdata (name, number PRIMARY KEY, grade, hours int, csa)''')
c.execute('''CREATE TABLE IF NOT EXISTS recorddata (timestamp timestamp, name, number, hours int)''')
#faker library for simulating data
fake = Faker()

class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'CSA Program Chapter Tracker'
        self.left = 50
        self.top = 75
        self.width = 1350
        self.height = 700
        self.init_UI()
        
    def init_UI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.create_student_table()
        self.create_record_table()
        self.recordTableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.create_calendar()

        # l is label
        self.lname = QLabel("Name", self)
        self.lnumber = QLabel("Student ID", self)
        self.lgrade = QLabel("Grade", self)
        self.lhours = QLabel("Hours", self)
        self.lcsa = QLabel("CSA", self)
        self.lstart = QLabel(self.my_calendar.selectedDate().toString(), self)
        self.lend = QLabel(self.my_calendar.selectedDate().toString(), self)

        self.lname.move(102.5, 540)
        self.lnumber.move(252.5, 540)
        self.lgrade.move(425, 540)
        self.lhours.move(562.5, 540)
        self.lcsa.move(710, 540)
        self.lstart.move(1025, 350)
        self.lend.move(1180, 350)

        self.lstart.setFrameStyle(QFrame.Panel | QFrame.Raised)
        self.lend.setFrameStyle(QFrame.Panel | QFrame.Raised)

        self.lstart.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.lstart.setAlignment(Qt.AlignCenter)
        self.lend.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.lend.setAlignment(Qt.AlignCenter)
        self.lstart.resize(self.lstart.sizeHint().width()+6, self.lstart.sizeHint().height()+3)
        self.lend.resize(self.lend.sizeHint().width()+6, self.lend.sizeHint().height()+3)

        # tb is textbox
        self.tbname = QLineEdit(self)
        self.tbnumber = QLineEdit(self)
        self.tbhours = QLineEdit(self)

        # b is combobox
        self.bgrade = QComboBox(self)
        self.bcsa = QComboBox(self)
        self.btimespan = QComboBox(self)

        self.bgrade.addItem("9th")
        self.bgrade.addItem("10th")
        self.bgrade.addItem("11th")
        self.bgrade.addItem("12th")

        self.bcsa.addItem("N/A")
        self.bcsa.addItem("CSA Community")
        self.bcsa.addItem("CSA Service")
        self.bcsa.addItem("CSA Achievement")

        self.btimespan.addItem("Weeks")
        self.btimespan.addItem("Months")
        self.btimespan.addItem("Custom")

        self.bgrade.move(368, 570)
        self.bcsa.move(652.5, 570)
        self.btimespan.move(815, 105)

        self.tbname.move(50, 570)
        self.tbnumber.move(200, 570)
        self.tbhours.move(510, 570)

        self.tbname.resize(140, 25)
        self.tbnumber.resize(165, 25)
        self.tbhours.resize(140, 25)
        
        self.bgrade.resize(140, 25)
        self.bcsa.resize(155, 25)

        self.bgrade.currentText()

        self.add_button = QPushButton('Add', self)
        self.add_button.setToolTip('Add student to database with given data')
        self.add_button.move(810, 567.5)
        self.add_button.clicked.connect(self.add_student)

        self.delete_button = QPushButton('Delete', self)
        self.delete_button.setToolTip('Delete selected rows from database')
        self.delete_button.move(810, 245)
        self.delete_button.clicked.connect(self.remove_students)

        self.print_button = QPushButton('Compile Data', self)
        self.print_button.setToolTip('Compile the current student data as a pdf')
        self.print_button.move(810, 65)
        self.print_button.resize(self.print_button.sizeHint())
        self.print_button.clicked.connect(self.compile_pdf)

        self.generate_button = QPushButton('Generate Student', self)
        self.generate_button.setToolTip('Generate a random student')
        self.generate_button.move(810, 605)
        self.generate_button.resize(self.generate_button.sizeHint())
        self.generate_button.clicked.connect(self.generate_students)

        self.start_date_button = QPushButton('Set Start', self)
        self.start_date_button.setToolTip('Set the start date for the data to compile at')
        self.start_date_button.move(1030, 315)
        self.start_date_button.resize(self.start_date_button.sizeHint())
        self.start_date_button.clicked.connect(self.set_start_date)

        self.end_date_button = QPushButton('Set End', self)
        self.end_date_button.setToolTip('Set the end date for the data to compile to')
        self.end_date_button.move(1185, 315)
        self.end_date_button.resize(self.end_date_button.sizeHint())
        self.end_date_button.clicked.connect(self.set_end_date)

        self.btimespan.resize(self.print_button.sizeHint())

        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(50, 70, 550, 200)

        self.layout.addWidget(self.recordTableWidget)
        self.layout.addWidget(self.studentTableWidget)
        self.layout.setSpacing(29)
        self.layout.setStretch(0, 3)
        self.layout.setStretch(1, 5)

        self.setLayout(self.layout)
        self.edit = True

        self.show()

    def create_student_table(self):
        self.studentTableWidget = QTableWidget()

        row_count = c.execute("SELECT COUNT(*) FROM studentdata").fetchone()[0]
        self.studentTableWidget.setRowCount(row_count)
        self.studentTableWidget.setColumnCount(5)
        self.studentTableWidget.setShowGrid(True)
        self.studentTableWidget.move(0,0)
        self.studentTableWidget.setHorizontalHeaderLabels(['Name', 'Student ID', 'Grade Level', 'Hours', 'CSA'])
        
        self.header = self.studentTableWidget.horizontalHeader()
        self.header.setSectionResizeMode(0, QHeaderView.Stretch)
        self.header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.header.setSectionResizeMode(4, QHeaderView.ResizeToContents)

        self.populate_student_data()
        self.studentTableWidget.itemChanged.connect(self.log_change)

    def create_record_table(self):
        self.recordTableWidget = QTableWidget()

        row_count = c.execute("SELECT COUNT(*) FROM recorddata").fetchone()[0]
        self.recordTableWidget.setRowCount(row_count)
        self.recordTableWidget.setColumnCount(4)
        self.recordTableWidget.setShowGrid(True)
        self.recordTableWidget.setHorizontalHeaderLabels(['Date', 'Name', 'Student ID', 'Hours Added'])
        
        self.header = self.recordTableWidget.horizontalHeader()
        self.header.setSectionResizeMode(0, QHeaderView.Stretch)
        self.header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.header.setSectionResizeMode(2, QHeaderView.ResizeToContents)

        self.populate_record_data()

    def create_calendar(self):
        self.my_calendar = QCalendarWidget(self)
        self.my_calendar.move(1000, 70)
        self.my_calendar.resize(300, 200)
        self.ldate = QLabel(self)
        self.ldate.move(1100, 280)
        self.ldate.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        self.ldate.setLineWidth(2)
        self.my_calendar.clicked[QDate].connect(self.on_date_selection)

        self.ldate.setText(self.my_calendar.selectedDate().toString())
        self.ldate.resize(self.ldate.sizeHint())

    def on_date_selection(self, date):
        self.ldate.setText(date.toString())
        self.ldate.resize(self.ldate.sizeHint())

    def populate_student_data(self):
        c.execute("SELECT * FROM studentdata")
        student_data = c.fetchall()
        self.edit = False
        self.studentTableWidget.setRowCount(0);
        self.studentTableWidget.setRowCount(len(student_data));
        for i in range(len(student_data)):
            self.studentTableWidget.setItem(i, 0, QTableWidgetItem(str(student_data[i][0])))
            self.studentTableWidget.setItem(i, 1, QTableWidgetItem(str(student_data[i][1])))
            self.studentTableWidget.setItem(i, 2, QTableWidgetItem(str(student_data[i][2])))
            self.studentTableWidget.setItem(i, 3, QTableWidgetItem(str(student_data[i][3])))
            self.studentTableWidget.setItem(i, 4, QTableWidgetItem(str(student_data[i][4])))
        self.edit = True

    def populate_record_data(self):
        c.execute("SELECT * FROM recorddata")
        record_data = c.fetchall()[::-1]
        self.recordTableWidget.setRowCount(0);
        self.recordTableWidget.setRowCount(len(record_data));
        for i in range(len(record_data)):
            self.recordTableWidget.setItem(i, 0, QTableWidgetItem(str(record_data[i][0])))
            self.recordTableWidget.setItem(i, 1, QTableWidgetItem(str(record_data[i][1])))
            self.recordTableWidget.setItem(i, 2, QTableWidgetItem(str(record_data[i][2])))
            self.recordTableWidget.setItem(i, 3, QTableWidgetItem(str(record_data[i][3])))

    def log_change(self, item):
        if self.edit:
            cols = {0:"name", 1:"number", 2:"grade", 3:"hours", 4:"csa"}
            c.execute("SELECT * FROM studentdata")
            student_data = c.fetchall()
            student_number = student_data[item.row()][1]
            student_name = student_data[item.row()][0]
            old_hours = student_data[item.row()][3]

            if item.column() == 3:
                hours = item.text()
                if hours.isdigit() == False:
                    try:
                        hours = eval(hours)
                        c.execute("UPDATE studentdata SET hours = ? WHERE number = ?", (hours, student_number,))
                        self.update_csa(float(hours), student_number, item)

                        c.execute("INSERT INTO recorddata VALUES (?, ?, ?, ?)", (datetime.now().strftime("%B %d, %Y %I:%M %p"), \
                        student_name, student_number, str('%.3f' % (float(hours)-old_hours)).rstrip('0').rstrip('.'),))
                        self.populate_record_data()
                    except Exception as e:
                        value_error = QMessageBox()
                        value_error.setIcon(QMessageBox.Critical)
                        value_error.setInformativeText('Please enter a valid expression or number')
                        value_error.setWindowTitle("Error: Value")
                        value_error.exec_()
                else:
                    c.execute("UPDATE studentdata SET hours = ? WHERE number = ?", (hours, student_number,))
                    self.update_csa(float(hours), student_number, item)

                    c.execute("INSERT INTO recorddata VALUES (?, ?, ?, ?)", (datetime.now().strftime("%B %d, %Y %I:%M %p"), \
                    student_name, student_number, str('%.3f' % (float(hours)-old_hours)).rstrip('0').rstrip('.'),))
                    self.populate_record_data()


            else:
                try:
                    c.execute("UPDATE studentdata SET {} = ? WHERE number = ?".format(cols[item.column()]), (item.text(), student_number,))
                    if item.column() == 0 or item.column() == 1:
                        c.execute("UPDATE recorddata SET {} = ? WHERE number = ?".format(cols[item.column()]), (item.text(), student_number,))
                        self.populate_record_data()
                except Exception as e:
                    insert_error = QMessageBox()
                    insert_error.setIcon(QMessageBox.Critical)
                    insert_error.setInformativeText('Please check your Student ID is unique')
                    insert_error.setWindowTitle("Error: Values")
                    insert_error.exec_()
                    replace_list = ['/'*(i+1) for i in range(len(student_data))]
                    replace_existing = [i[1] for i in student_data if i[1] in replace_list]
                    replace_new = [i for i in replace_list if i not in replace_existing]
                    replace_new.sort(key=len)
                    c.execute("UPDATE studentdata SET number = ? WHERE number = ?", (replace_new[0], student_number,))
                    c.execute("UPDATE recorddata SET number = ? WHERE number = ?", (replace_new[0], student_number,))
                    self.studentTableWidget.setItem(item.row(), item.column(), QTableWidgetItem(replace_new[0]))


            conn.commit()

    def update_csa(self, hours, student_number, item):
        csa_switcher = {
            50 : "CSA Community",
            200 : "CSA Service",
            500 : "CSA Achievement"
        }
        adjusted_hours = 500 if hours >= 500 else 200 if hours >= 200 else 50 if hours >= 50 else 0
        csa_award = csa_switcher.get(adjusted_hours, "N/A")
        c.execute("UPDATE studentdata SET csa = ? WHERE number = ?", (csa_award, student_number))
        conn.commit()
        
        try:
            self.studentTableWidget.blockSignals(True)
            self.studentTableWidget.setItem(item.row(), 4, QTableWidgetItem(str(csa_award)))
            self.studentTableWidget.setItem(item.row(), 3, QTableWidgetItem(str('%.3f' % hours).rstrip('0').rstrip('.')))
            self.studentTableWidget.blockSignals(False)
        except Exception as e:
            print(e)

        c.execute("SELECT * FROM studentdata")

    @pyqtSlot()
    def add_student(self):
        # i is input
        iname = self.tbname.text()
        inumber = self.tbnumber.text()
        igrade = str(self.bgrade.currentText())
        ihours = self.tbhours.text()
        icsa = str(self.bcsa.currentText())
        try:
            for i in range(5):
                if [iname, inumber, igrade, ihours, icsa][i] == '':
                    raise Exception
            int(ihours)
            c.execute('''INSERT INTO studentdata VALUES (?, ?, ?, ?, ?)''', (iname, inumber, igrade, ihours, icsa,))
            c.execute('''INSERT INTO recorddata VALUES (?, ?, ?, ?)''', (datetime.now().strftime("%B %d, %Y %I:%M %p"), iname, inumber, ihours,))
            conn.commit()

            self.populate_student_data()
            self.populate_record_data()
        except Exception as e:

            insert_error = QMessageBox()
            insert_error.setIcon(QMessageBox.Critical)
            insert_error.setInformativeText('Please check your Student ID is unique, all inputs are filled out, and hours is a number')
            insert_error.setWindowTitle("Error: Values")
            insert_error.exec_()

    @pyqtSlot()
    def remove_students(self):
        selected = self.studentTableWidget.selectedItems()
        selected_dict = {}
        for item in selected:
            if item.row() in selected_dict:
                selected_dict[item.row()] += 1
            else:
                selected_dict[item.row()] = 1

        c.execute('''SELECT * FROM studentdata''')
        student_data = c.fetchall()
        numbers = ()
        
        for key in reversed(list(selected_dict.keys())):
            if selected_dict[key] == 5:
                self.studentTableWidget.removeRow(key)
                numbers += (student_data[key][1],)

        c.execute("DELETE FROM studentdata WHERE number IN ({})".format(", ".join("?"*len(numbers))), numbers)
        c.execute("DELETE FROM recorddata WHERE number IN ({})".format(", ".join("?"*len(numbers))), numbers)
        self.populate_record_data()

        conn.commit()
    
    @pyqtSlot()
    def compile_pdf(self):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("courier", size=12)

        
        if self.btimespan.currentText() == "Weeks":
            start_range = (datetime.now() - timedelta(7))
            end_range = datetime.now()
            txt = "Last Week Report"
        elif self.btimespan.currentText() == "Months":
            start_range = (datetime.now() - timedelta(30))
            end_range = datetime.now()
            txt = "Last Week Report"
        else:
            start_range = datetime.strptime(self.lstart.text(), "%a %b %d %Y")
            end_range = datetime.strptime(self.lend.text(), "%a %b %d %Y")
            txt = "Custom Report: %s - %s"%(start_range.strftime("%B %d, %Y"), end_range.strftime("%B %d, %Y"))

        pdf.cell(200, 10, txt=txt, ln=1, align="C")
        c.execute('''SELECT * FROM recorddata''')
        datafetch = c.fetchall()
        record_data = [record for record in datafetch if datetime.strptime(record[0], ("%B %d, %Y %I:%M %p")) >= start_range and\
             datetime.strptime(record[0], ("%B %d, %Y %I:%M %p")) <= end_range][::-1]

        if record_data == []:
            insert_error = QMessageBox()
            insert_error.setIcon(QMessageBox.Critical)
            insert_error.setInformativeText('No data to Compile')
            insert_error.setWindowTitle("Error: Data")
            insert_error.exec_()
        else:
            field1_max = len(max([str(record[0]) for record in record_data], key=len))
            field2_max = len(max([str(record[1]) for record in record_data], key=len))
            field3_max = len(max([str(record[2]) for record in record_data], key=len))
            field4_max = len(max([str(record[3]) for record in record_data], key=len))

            field_maxes = [field1_max, field2_max, field3_max, field4_max]
            xpos = 200
            ypos = 15
            yline = 35
            iconst = 2.725
            sconst = 2.4
            spacer = 0 if sum(field_maxes) > 75 else int((75-sum(field_maxes))/3)
            line_count = 0
            page_lines = 17

            if spacer == 0:
                pdf.set_font("courier", size=9)
                spacer = 0 if sum(field_maxes) > 100 else int((100-sum(field_maxes))/3)
                iconst *= 0.75
                sconst *= 0.75
                
            for record in record_data:
                txt = ''
                for i in range(len(list(record))):
                    txt += str(record[i]) + ' ' * int(field_maxes[i]-len(str(record[i])) + spacer)
                
                pdf.cell(xpos, ypos, txt=txt, ln=1, align="L")
                pdf.line(0, yline, 400, yline)
                yline += 15
                line_count += 1
                
                newx = 0
                if line_count % page_lines == 0:
                    yline = 10
                    title_space = 25 if line_count == page_lines else 10
                    for i in field_maxes[0:3]:
                        newx += (i) * iconst + (spacer) * sconst
                        pdf.line(newx, title_space, newx, 20 + 15 * page_lines)
                elif line_count == len(record_data):
                    title_space = 25 if line_count < page_lines else 10
                    footer_space = 20 if line_count < page_lines else 10
                    for i in field_maxes[0:3]:
                        newx += (i) * iconst + (spacer) * sconst
                        pdf.line(newx, title_space, newx, footer_space + 15 * (line_count%page_lines))
            
            if len(record_data) > page_lines:
                pdf.line(0, yline, 400, yline)

            c.execute('''SELECT name, number FROM studentdata''')
            student_data = c.fetchall()
            for student in student_data:
                if student[1] in [record[2] for record in record_data]:
                    total_user_hours = str(round(sum([record[3] for record in record_data if record[2] == student[1]]), 3))
                    txt = ' ' * int(sum(field_maxes) + spacer*3 - len(total_user_hours) - len(student[0]) - 2) + student[0] + ': ' + total_user_hours
                    pdf.cell(xpos, ypos, txt=txt, ln=1, align="L")

            total_hours = str(round(sum([record[3] for record in record_data]), 3))
            txt = ' ' * int(sum(field_maxes) + spacer*3 - len(total_hours) - 13) + "Total Hours: " + total_hours

            pdf.cell(xpos, ypos, txt=txt, ln=1, align="L")


            pdf.output("community_service_program.pdf")

    #generate random student with hours added on a random date
    @pyqtSlot()
    def generate_students(self):
        start_dt = date.today().replace(day=1, month=1).toordinal()
        end_dt = date.today().toordinal()
        random_day = date.fromordinal(random.randint(start_dt, end_dt)).strftime("%B %d, %Y %I:%M %p")

        random_name = fake.name()
        random_id = fake.ssn()
        random_grade = random.choice(['9th', '10th', '11th', '12th'])
        random_hours = random.randint(1, 500)
        csa_award = 'CSA Community' if random_hours >= 500 else ('CSA Service' if random_hours >= 200 else ('CSA Achievement' if random_hours >= 50 else 'N/A'))
        c.execute('''INSERT INTO studentdata VALUES (?, ?, ?, ?, ?)''', (random_name, str(random_id), random_grade, random_hours, csa_award,))
        c.execute('''INSERT INTO recorddata VALUES (?, ?, ?, ?)''', (str(random_day), random_name, str(random_id), random_hours,))
        conn.commit()

        self.populate_record_data()
        self.populate_student_data()

    @pyqtSlot()
    def set_start_date(self):
        if datetime.strptime(self.lend.text(), "%a %b %d %Y") < datetime.strptime(self.my_calendar.selectedDate().toString(), "%a %b %d %Y"):
            self.lend.setText(self.my_calendar.selectedDate().toString())
            self.lend.resize(self.lend.sizeHint())
        self.lstart.setText(self.my_calendar.selectedDate().toString())
        self.lstart.resize(self.lstart.sizeHint().width()+6, self.lstart.sizeHint().height()+3)


    @pyqtSlot()
    def set_end_date(self):
        if datetime.strptime(self.my_calendar.selectedDate().toString(), "%a %b %d %Y") < datetime.strptime(self.lstart.text(), "%a %b %d %Y"):
            date_error = QMessageBox()
            date_error.setIcon(QMessageBox.Critical)
            date_error.setInformativeText('End date cannot be before start date')
            date_error.setWindowTitle("Error: Date")
            date_error.exec_()
        else:
            self.lend.setText(self.my_calendar.selectedDate().toString())
            self.lend.resize(self.lend.sizeHint().width()+6, self.lend.sizeHint().height()+3)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())