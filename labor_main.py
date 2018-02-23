import sys
import labordialog
import os
import pyodbc
import PyQt5
import datetime
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QPushButton, QAction, QLineEdit, QMessageBox, QLCDNumber
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import pyqtSlot, QTimer, QTime

class LaborTracker(QMainWindow, labordialog.Ui_MainWindow):

	def __init__(self, parent=None):

		super(LaborTracker, self).__init__(parent)
		self.setupUi(self)

		# Set focus to first textbox on load, so the WOID can be scanned.
		self.txtEID.setFocus()

		# Call method to disable buttons.
		self.disable_start()

		# Call method after scan on WOID.
		self.txtWOID.returnPressed.connect(self.scan_work_order)

		# Call method after scan on Employee Clock #.
		self.txtEID.returnPressed.connect(self.scan_employee)

		# Call method to handle Start btn.
		self.btnStart.clicked.connect(self.start_labor)

		# Call method to handle Stop btn.
		self.btnStop.clicked.connect(self.stop_labor)

		# Call method to clear form.
		self.btnClear.clicked.connect(self.clear_form)

		# Define message to be passed to message box class.
		LaborTracker.message = ''

		self.call_clock()

	# Method for connecting through ODBC to SQL Server.
	def connect(self):

		global cnxn, cursor

		cnxn = pyodbc.connect("DSN=sqlserver;UID=FSUser;PWD=1@freeman")
		cursor = cnxn.cursor()

	def disconnect(self):
		cursor = cnxn.cursor()
		cursor.close()
		del cursor
		cnxn.close()

	# Method for scanning WOID.
	def scan_work_order(self):

		global WOID, WOName

		WOID = self.txtWOID.text()

		self.connect()

		cursor.execute("SELECT [Work Order].[WOID], [Work Order].[WO Status], "
					   "[Products].[Product Code] AS Name "
					   "FROM [Work Order] INNER JOIN [Products] ON "
					   "[Work Order].[ProductID] = [Products].[ID] "
					   "WHERE ([Work Order].[WO Status] = 2 OR [Work Order].[WO Status] = 3) "
					   "AND ([Work Order].[WOID] = ?)", (WOID))

		row = cursor.fetchone()
		# If we have data. Any Work Order that is RELASED or WIP.
		if row:

			self.btnStart.setFocus()
			WOName = row.Name

		else:
			# If the above criteria is not met.
			self.txtWOID.clear()
			self.txtWOID.setFocus()
			return

	# Method for scanning employee Clock #.
	def scan_employee(self):

		global ClockID, TimeNow, RecordID, TimeIn, LastLaborCode, CurrentWOID

		TimeNow = datetime.datetime.now()
		ClockID = self.txtEID.text()

		if self.get_employee_info() is False:
			self.clear_form()
			msgBox = QMessageBox.about(self, "Error", \
				"You must be a current employee on the labor list to record labor.")
			return

		self.connect()
		cursor.execute("SELECT ID, EmpClockID, WOID, TimeIn, TimeOut, LaborCode "
					   "FROM LaborTest WHERE EmpClockID = ? "
					   "ORDER BY ID DESC", (EmpID))

		row = cursor.fetchone()

		if row:
			# If the employee has a record for labor. All will have this except new hire on first day.
			RecordID = row.ID
			# Checking for a completed record, i.e. both Time In and Time Out fields populated.
			if (row.TimeIn is not None and row.TimeOut is not None) \
			or (row.TimeIn is None and row.TimeOut is None):
				# If so, they're starting new work. So we call function that enables starting.
				LastLaborCode = row.LaborCode
				self.auto_select_labor_code()
				self.enable_start()
				self.txtWOID.setFocus()
			else:
				# This means that they're currently working, so they'll only be able to stop work.
				TimeIn = row.TimeIn
				self.get_work_time()
				CurrentWOID = row.WOID
				self.get_woid_name()
				self.btnStop.setEnabled(True)
				self.btnStop.setFocus()

		else:
			# This is to handle new hires on their first day.
			# We could remove this if we automatically add an inital
			# entry with all fields completed.
			self.enable_start()
			self.txtWOID.setFocus()



	# Method for returning labor code from radio button selection.
	def get_labor_code(self):

		global LaborCode

		if self.rbFab.isChecked():
			LaborCode = 1
		elif self.rbWeld.isChecked():
			LaborCode = 2
		elif self.rbAssembly.isChecked():
			LaborCode = 3
		elif self.rbPaint.isChecked():
			LaborCode = 4
		elif self.rbFinal.isChecked():
			LaborCode = 5
		elif self.rbIndirect.isChecked():
			LaborCode = 6
		elif self.rbElectric.isChecked():
			LaborCode = 10
		elif self.rbShipReceive.isChecked():
			LaborCode = 11
		elif self.rbMaterialHandling.isChecked():
			LaborCode = 13
		elif self.rbLab.isChecked():
			LaborCode = 20
		else:
			return False

	# Method for returning name of employee to be used in messagebox 
	# upon completion of start/stop methods.
	def get_employee_info(self):

		global FirstName, EmpID

		self.connect()
		cursor.execute("SELECT [ID], [First Name] AS FirstName, ClockID, [Labor List], [Status ID] " 
					   "FROM Employees "
					   "WHERE [Labor List] = 1 AND [Status ID] = 1 AND ClockID = ?", (ClockID))

		row = cursor.fetchone()

		if row:
			FirstName = row.FirstName
			EmpID = row.ID
		else:
			return False

	# Method for starting labor.
	def start_labor(self):

		if self.validate_form() is False:
			return
			
		self.connect()
		cursor.execute("INSERT INTO LaborTest (EmpClockID, WOID, TimeIn, LaborCode) "
							   "VALUES (?, ?, ?, ?)", (EmpID, WOID, TimeNow, LaborCode))
		cnxn.commit()

		self.get_employee_info()

		LaborTracker.message = "Thanks " + FirstName + "! You've begun working on " + WOName + " - " + WOID + "."

		self.clear_form()
		self.call_msg_timer()

	# Method for stopping labor, fewer controls / options.
	def stop_labor(self):

		self.connect()
		cursor.execute("UPDATE LaborTest SET TimeOut = ? WHERE ID = ?", (TimeNow, RecordID))
		cnxn.commit()

		self.get_employee_info()

		LaborTracker.message = ("Thanks " + FirstName +
								   "! You've stopped working on " + CurrentName + 
								   " - " + str(CurrentWOID) +
								   ". Your work segment was: " + tDelta + ".")

		self.clear_form()
		self.call_msg_timer()

	# Form validation to confirm there is data, meant to handle potential errors.
	def validate_form(self):

		if self.txtWOID.text() == "":
			self.txtWOID.setFocus()
			return False

		if self.txtEID.text() == "":
			self.txtEID.setFocus()
			return False

		if self.get_labor_code() == False:
			return False

	def get_woid_name(self):

		global CurrentName
		self.connect()
		cursor.execute("SELECT [Work Order].WOID, Products.[Product Code] AS Name FROM "
					   "[Work Order] INNER JOIN Products ON "
					   "[Work Order].ProductID = Products.ID "
					   "WHERE WOID = ?", (CurrentWOID))
		row = cursor.fetchone()

		if row:
			CurrentName = row.Name
		else:
			return

	# Method for enabling radio buttons and start button.
	def enable_start(self):

		self.rbFab.setEnabled(True)
		self.rbWeld.setEnabled(True)
		self.rbAssembly.setEnabled(True)
		self.rbPaint.setEnabled(True)
		self.rbFinal.setEnabled(True)
		self.rbIndirect.setEnabled(True)
		self.rbElectric.setEnabled(True)
		self.rbShipReceive.setEnabled(True)
		self.rbMaterialHandling.setEnabled(True)
		self.rbLab.setEnabled(True)
		self.btnStart.setEnabled(True)

	# On load method and after update on start/stop so that no
	# buttons can be pressed.
	def disable_start(self):

		self.rbFab.setEnabled(False)
		self.rbWeld.setEnabled(False)
		self.rbAssembly.setEnabled(False)
		self.rbPaint.setEnabled(False)
		self.rbFinal.setEnabled(False)
		self.rbIndirect.setEnabled(False)
		self.rbElectric.setEnabled(False)
		self.rbShipReceive.setEnabled(False)
		self.rbMaterialHandling.setEnabled(False)
		self.rbLab.setEnabled(False)
		self.btnStart.setEnabled(False)
		self.btnStop.setEnabled(False)

	# Clear form method after updating or inserting a record.
	def clear_form(self):

		self.txtWOID.clear()
		self.txtEID.clear()
		self.txtEID.setFocus()
		self.btnStart.setEnabled(False)
		self.btnStop.setEnabled(False)
		self.disable_start()

	# Return time spent on work segment to employee.
	def get_work_time(self):

		global tDelta

		FMT = '%H:%M:%S'
		strTimeIn = TimeIn.strftime(FMT)
		strTimeNow = TimeNow.strftime(FMT)

		tDelta = str(datetime.datetime.strptime(strTimeNow, FMT) - \
		datetime.datetime.strptime(strTimeIn, FMT))

	def auto_select_labor_code(self):

		if LastLaborCode == 1:
			self.rbFab.setChecked(True)
		elif LastLaborCode == 2:
			self.rbWeld.setChecked(True)
		elif LastLaborCode == 3:
			self.rbAssembly.setChecked(True)
		elif LastLaborCode == 4:
			self.rbPaint.setChecked(True)
		elif LastLaborCode == 5:
			self.rbFinal.setChecked(True)
		elif LastLaborCode == 6:
			self.rbIndirect.setChecked(True)
		elif LastLaborCode == 10:
			self.rbElectric.setChecked(True)
		elif LastLaborCode == 11:
			self.rbShipReceive.setChecked(True)
		elif LastLaborCode == 13:
			self.rbMaterialHandling.setChecked(True)
		elif LastLaborCode == 20:
			self.rbLab.setChecked(True)
		else:
			return

	def call_msg_timer(self):
		msgBox = TimerMessageBox(10, self)
		msgBox.exec_()

	def call_clock(self):
		# x = DigitalClock(self)
		# x.show()
		self.clock()

	def clock(self):
		timer = QTimer(self)
		timer.timeout.connect(self.showTime)
		timer.start(1000)
		self.showTime()

	def showTime(self):
		time = QTime.currentTime()
		text = time.toString('hh:mm')
		if (time.second() % 2) == 0:
			text = text[:2] + ' ' + text[3:]
			
		self.lcdTime.display(text)

# Class for automatic close on messagebox after start and stop work.
# This closes after 10 seconds but can be closed manually by scanning.
class TimerMessageBox(QMessageBox):

	def __init__(self, timeout=10, parent=None):
		super(TimerMessageBox, self).__init__(parent)
		self.setWindowTitle('Clock Successful')
		self.time_to_wait = timeout
		font = QFont()
		font.setPointSize(18)
		self.setFont(font)
		self.setText(LaborTracker.message)
		self.setStandardButtons(QMessageBox.Ok)
		self.timer = QTimer(self)
		self.timer.setInterval(1000)
		self.timer.timeout.connect(self.change_timer)
		self.change_timer()
		self.timer.start()

	def change_timer(self):
		self.time_to_wait -= 1
		if self.time_to_wait <= 0:
			self.close()

	def closeEvent(self, event):
		self.timer.stop()
		event.accept()

def main():

	app = QApplication(sys.argv)
	form = LaborTracker()
	form.show()
	app.exec_()

if __name__=='__main__':

	main()