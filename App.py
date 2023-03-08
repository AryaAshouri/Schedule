from PyQt5.uic import loadUi
from PyQt5.QtWidgets import * 
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import * 
from PyQt5.QtCore import *
import sys, sqlite3, datetime, xlwt, time
from datetime import datetime
from xlwt import Workbook
from selenium import webdriver
from time import sleep

class Main(QMainWindow):
	def __init__(self):
		super(Main, self).__init__()
		loadUi("static/ui/UI.ui", self)
		self.setWindowTitle("Schedule")
		self.setWindowIcon(QtGui.QIcon('static/img/icon.png'))
		self.label_7.setVisible(False)

		self.add_note.clicked.connect(self.add_another_note)
		self.all_notes.clicked.connect(self.all_notes_content)
		self.back_1.clicked.connect(self.back_notes_content1)
		self.back_9.clicked.connect(self.send_schedule_to_bale)
		self.back_5.clicked.connect(self.download_schedule)
		self.back_8.clicked.connect(self.go_to_bale)
		self.back_6.clicked.connect(self.back_to_all)
		self.back_4.clicked.connect(self.delete_item)
		self.back_2.clicked.connect(self.go_down)
		self.back_7.clicked.connect(self.sec_code)
		self.back_3.clicked.connect(self.go_up)

	def back_notes_content1(self):
		self.stackedWidget.setCurrentIndex(0)
		self.content.clear()
		self.content_2.clear()
		self.label_2.clear()
		self.label_5.clear()
		self.label_3.clear()
		self.label_6.clear()
		self.content_3.clear()
		self.content_4.clear()
		self.listWidget.clear()

	def add_another_note(self):

		content_of_note = self.content.text()
		content_of_time = self.content_2.text()
		now = datetime.now()

		if (content_of_note != "" and content_of_time != ""):
			self.label_2.setText("محتوا ذخیره شد")

			connection = sqlite3.connect("Database/Data.db")
			cursor = connection.cursor()

			cursor.execute('''CREATE TABLE IF NOT EXISTS Note(Content TEXT, Period TEXT, Time TEXT, Date TEXT)''')
			cursor.execute("INSERT INTO Note VALUES(?, ?, ?, ?)", (content_of_note, content_of_time, now.strftime("%H:%M:%S"), now.strftime("%Y:%m:%d")))
			connection.commit()
			connection.close()

	def all_notes_content(self):
		self.stackedWidget.setCurrentIndex(1)
		self.content.setText = ""

		connection = sqlite3.connect("Database/Data.db")
		cursor = connection.cursor()
		tasks = cursor.execute('''SELECT Content From Note''')
		item_text_list = []

		for i in tasks:
			item_text_list.append(i[0])

		for item_text in item_text_list:
			item = QListWidgetItem(item_text)
			item.setTextAlignment(Qt.AlignHCenter)
			self.listWidget.addItem(item)

		connection.commit()
		connection.close()

	def go_down(self):
		index = self.listWidget.currentRow()
		if (index >= 0):
			item = self.listWidget.takeItem(index)
			self.listWidget.insertItem(index+1, item)
			self.listWidget.setCurrentItem(item)

	def go_up(self):
		index = self.listWidget.currentRow()
		if (index >= 0):
			item = self.listWidget.takeItem(index)
			self.listWidget.insertItem(index-1, item)
			self.listWidget.setCurrentItem(item)

	def delete_item(self):
		connection = sqlite3.connect("Database/Data.db")
		cursor = connection.cursor()
		tasks = cursor.execute('''SELECT Content From Note''')
		counter = 0
		for i in tasks:
			counter += 1

		if (counter >= 1):
			connection = sqlite3.connect("Database/Data.db")
			cursor = connection.cursor()

			try:
				index = self.listWidget.currentRow()
				task_item_name = self.listWidget.takeItem(index).text()
				cursor.execute("DELETE FROM Note WHERE Content = ?", (task_item_name, ))
				connection.commit()
				connection.close()
				self.label_5.setText("مورد حذف شد")

				wb = Workbook()
				sheet1 = wb.add_sheet("Sheet 1")
				sheet1.write(0, 0, "Content")
				sheet1.write(0, 1, "Period")
				sheet1.write(0, 2, "Date")
				sheet1.write(0, 3, "Time")

				connection = sqlite3.connect("Database/Data.db")
				cursor = connection.cursor()
				content = cursor.execute('''SELECT Content From Note''')
				content = cursor.fetchall()
				period = cursor.execute('''SELECT Period From Note''')
				period = cursor.fetchall()
				date = cursor.execute('''SELECT Date From Note''')
				date = cursor.fetchall()
				time = cursor.execute('''SELECT Time From Note''')
				time = cursor.fetchall()

				for i in range(0, len(content)):
					sheet1.write(i+1, 0, content[i])
				for i in range(0, len(period)):
					sheet1.write(i+1, 1, period[i])
				for i in range(0, len(date)):
					sheet1.write(i+1, 2, date[i])
				for i in range(0, len(time)):
					sheet1.write(i+1, 3, time[i])

				connection.commit()
				connection.close()
				wb.save("Excel/Schedule.xls")

			except:
				pass

	def download_schedule(self):
		connection = sqlite3.connect("Database/Data.db")
		cursor = connection.cursor()
		tasks = cursor.execute('''SELECT Content From Note''')
		counter = 0

		for i in tasks:
			counter += 1

		if (counter >= 1):
			wb = Workbook()
			sheet1 = wb.add_sheet("Sheet 1")
			sheet1.write(0, 0, "Content")
			sheet1.write(0, 1, "Period")
			sheet1.write(0, 2, "Date")
			sheet1.write(0, 3, "Time")

			connection = sqlite3.connect("Database/Data.db")
			cursor = connection.cursor()
			content = cursor.execute('''SELECT Content From Note''')
			content = cursor.fetchall()
			period = cursor.execute('''SELECT Period From Note''')
			period = cursor.fetchall()
			date = cursor.execute('''SELECT Date From Note''')
			date = cursor.fetchall()
			time = cursor.execute('''SELECT Time From Note''')
			time = cursor.fetchall()

			for i in range(0, len(content)):
				sheet1.write(i+1, 0, content[i])
			for i in range(0, len(period)):
				sheet1.write(i+1, 1, period[i])
			for i in range(0, len(date)):
				sheet1.write(i+1, 2, date[i])
			for i in range(0, len(time)):
				sheet1.write(i+1, 3, time[i])

			connection.commit()
			connection.close()
			wb.save("Excel/Schedule.xls")
			self.label_5.setText("برنامه ذخیر شد")

		else:
			self.label_5.setText("برنامه یافت نشد")


	def back_to_all(self):
		self.stackedWidget.setCurrentIndex(1)
		self.label_6.clear()
		self.content_4.clear()
		self.content_3.clear()
		self.label_5.clear()

	def send_schedule_to_bale(self):
		self.stackedWidget.setCurrentIndex(2)
		self.label_5.clear()

	def go_to_bale(self):
		self.label_6.clear()
		connection = sqlite3.connect("Database/Data.db")
		cursor = connection.cursor()
		tasks = cursor.execute('''SELECT Content From Note''')
		counter = 0

		for i in tasks:
			counter += 1

		if (counter > 0):
			if (self.content_3.text() != "" and self.content_4.text() != ""):
				driver = webdriver.Chrome("Driver/chromedriver")
				driver.maximize_window()
				driver.get("http://web.bale.ai")

				time.sleep(7)
				driver.find_element_by_xpath('//*[@id="root"]/div[1]/div/div/div[2]/button').click()
				time.sleep(2)
				driver.find_element_by_xpath('//*[@id="شماره همراه"]').send_keys(self.content_3.text())
				driver.find_element_by_xpath('//*[@id="root"]/div[1]/div/div/div/div[2]/button').click()
				time.sleep(2)
				driver.quit()

				self.content_3.setVisible(False)
				self.content_4.setVisible(False)
				self.back_6.setVisible(False)
				self.back_8.setVisible(False)
			else:
				self.label_6.setText("ورودی ها را پر کنید")
		else:
			self.label_6.setText("برنامه یافت نشد")

	def sec_code(self):
		if (self.content_5.text() != ""):
			driver = webdriver.Chrome("Driver/chromedriver")
			driver.maximize_window()
			driver.get("http://web.bale.ai")

			time.sleep(7)
			driver.find_element_by_xpath('//*[@id="root"]/div[1]/div/div/div[2]/button').click()
			time.sleep(2)
			driver.find_element_by_xpath('//*[@id="شماره همراه"]').send_keys(self.content_3.text())
			driver.find_element_by_xpath('//*[@id="root"]/div[1]/div/div/div/div[2]/button').click()
			time.sleep(2)
			driver.find_element_by_xpath('//*[@id="کد ورود"]').send_keys(self.content_5.text())
			driver.find_element_by_xpath('//*[@id="root"]/div[1]/div/div/div[2]/button').click()
			time.sleep(7)
			driver.find_element_by_xpath('/html/body/div[1]/div[1]/div/div/div[2]/div[1]/div/input').send_keys(self.content_4.text())
			time.sleep(360)

		else:
			self.label_7.setVisible(True)
			self.label_7.setText("ورودی ها را پر کنید")

main = QApplication(sys.argv)
app = Main()
app.show()
main.exec_()