from PyQt5 import  QtWidgets, QtGui, QtCore
import pandas as pd
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import time
import traceback, sys, os, random
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import smbus


class Window(QtWidgets.QWidget):
	def __init__(self, **kwargs):
		super(Window, self).__init__(**kwargs)
		
		self.setWindowTitle("Merici aplikace")
		
		self.graph()
		self.setFixedSize(900, 400)
        
		self.layout_form = QtWidgets.QHBoxLayout()
		self.setLayout(self.layout_form)
		
		self.graph_column_boxlayout = QtWidgets.QVBoxLayout()
		self.buttons_column_boxlayout = QtWidgets.QVBoxLayout()
		self.instantaneous_row_boxlayout = QtWidgets.QHBoxLayout()
		
		self.instantaneous_value_label = QtWidgets.QLabel("Okamzita hodnota:")
		self.instantaneous_value = QtWidgets.QLabel(" ")
		
		for x in (self.instantaneous_value_label, self.instantaneous_value):
			x.setFont(QFont('Arial', 15))
			self.instantaneous_row_boxlayout.addWidget(x)

		self.instantaneous_row_boxlayout.addStretch()
		
		self.graph_column_boxlayout.addLayout(self.instantaneous_row_boxlayout)
		self.graph_column_boxlayout.addWidget(self.canvas)		
		self.layout_form.addLayout(self.graph_column_boxlayout)
		
		self.connect_button = QtWidgets.QPushButton("Pripojit a merit")
		self.pause_button = QtWidgets.QPushButton("Pozastavit")
		self.toexcel_button = QtWidgets.QPushButton("Zapsat do excelu")
		self.delete_button = QtWidgets.QPushButton("Vymazat")
		
		for x in (self.connect_button, self.pause_button, self.toexcel_button, self.delete_button):
			x.setFixedSize(200, 90)
			x.setFont(QFont('Arial', 15))
			self.buttons_column_boxlayout.addWidget(x)

		self.layout_form.addLayout(self.buttons_column_boxlayout)
		
		buttons = [self.connect_button, self.pause_button, self.toexcel_button, self.delete_button]
		methods = [self.connect, self.pause, self.toexcel, self.delete]
		
		for button, method in zip(buttons,methods):
			button.clicked.connect(method)
	
		self.show()
		
	def setup(self):
		self.logic = apka.logic

		
	def graph(self):
		self.ydata = []
		self.xdata = []
		self.min_x = 0
		self.max_x = 1000
		self.min_y = -1
		self.max_y = 5
		
		self.figure = Figure(figsize = (5, 5), dpi = 100)
		self.ax = self.figure.add_subplot(1, 1, 1)
		self.line1, = self.ax.plot([], [])
		self.ax.set_xlim(self.min_x, self.max_x)
		self.ax.set_ylim(self.min_y, self.max_y)
		self.ax.set_title("Prubeh signalu")
		self.ax.grid()
		self.canvas = FigureCanvas(self.figure)
		
	def connect(self):
		self.logic.connect()

	
	def pause(self):
		self.logic.pause()

		
	def delete(self):
		self.logic.delete()

		
	def toexcel(self):
		self.logic.toexcel()

class Logic():
	
	def __init__(self):

		self.i = 0
		self.timer = QTimer()
		self.timer.setInterval(100)
		
		self.device_bus = 1
		self.device_addr = 0x48
		self.bus = smbus.SMBus(self.device_bus)
		self.control_byte = 0x00
		
	def setup(self):
		self.window = apka.window		

	def read_excel(self):
		
		try:
			
			self.i = self.i + 1
			
			self.bus.write_byte(self.device_addr, self.control_byte)
			data = self.bus.read_byte(self.device_addr)
			self.voltage = data*(3.3/255)
	
			self.window.instantaneous_value.setText(str(self.voltage))
	
			self.window.ydata.append(self.voltage)
			self.window.xdata.append(self.i)
	
			
			self.window.line1.set_xdata(self.window.xdata)
			self.window.line1.set_ydata(self.window.ydata)
			
			self.window.canvas.draw()
			self.window.figure.canvas.flush_events()

		except:
			self.pause()
			self.msg = QMessageBox()
			self.msg.setText("Nepripojeno.")
			self.msg.show()
			
	def connect(self):
		self.timer.timeout.connect(self.read_excel)
		self.timer.start()
		
	def pause(self):
		self.timer.stop()
		
	def delete(self):
		self.timer.stop()		
		self.window.canvas.close()
		self.i = 0
		self.ii = 0
		self.window.graph()
		self.window.graph_column_boxlayout.addWidget(self.window.canvas)
		self.window.instantaneous_value.setText(' ')
		
	def toexcel(self):  
	
		if self.timer.isActive:
			self.timer.stop()
			
		try:	
			dialog = QFileDialog()
			self.folder_path = dialog.getExistingDirectory(None, "Zvolte slozku")
			self.folder_path = self.folder_path + '/cisla.xlsx'
			
			zapis = pd.DataFrame(list(zip(self.window.xdata, self.window.ydata)), columns=['x', 'y'])	
			
			with pd.ExcelWriter(self.folder_path, mode= 'w', engine='xlsxwriter') as writer:
				zapis.to_excel(writer, sheet_name='Sheet1', index=False)
				
		except:
			self.msg = QMessageBox()
			self.msg.setText("Nezvolena cesta pro ulozeni excelu")
			self.msg.setStandardButtons(QMessageBox.Ok)
			self.msg.show()
	
class App(QtWidgets.QApplication):
	def __init__(self):
		super(App, self).__init__(sys.argv)
		
	def builder(self):
		self.window = Window()
		self.logic = Logic()
		self.window.setup()
		self.logic.setup()
		sys.exit(self.exec())
		
apka = App()
apka.builder()
