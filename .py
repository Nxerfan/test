import pandas as pd
import xlsxwriter
from PyQt6.QtWidgets import *
from PyQt6.QtCore import QSize
import sys

def user_input(value):
    DF = pd.read_excel(value, header=None)
    MyArray = []

    for i in DF:
        mylist = DF[i].tolist()
        MyArray.append(mylist)
    
    CountColumn=0
    Result=0
    for i in MyArray:
        CountColumn+=1
        for item in i:
            Result+=1
    Rows=int(Result/CountColumn)
    printer=f"Rows: {Rows}\nColumn: {CountColumn}\nItems: {Result}"
    

    return DF,printer



def process(data):
    MyArray = []

    for i in data:
        mylist = data[i].tolist()
        MyArray.append(mylist)
    print(MyArray)
    
    CountColumn=0
    CountRow=0
    reversed = []
    for i in MyArray:
        CountColumn+=1
        reversed_list = []
        for item in i:
            CountRow+=1
            reversed_list.append(item[::-1])
        reversed.append(reversed_list)


    return reversed


def output(entry,file):
    workbook = xlsxwriter.Workbook(file)
    worksheet = workbook.add_worksheet()

    for col_idx, column in enumerate(entry):
        for row_idx, name in enumerate(column):
            worksheet.write(row_idx, col_idx, name)

    workbook.close()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()  

        self.setWindowTitle("Parvizi_Sadeghi")
        self.OutPut=""
        self.InPut=""
        self.InPutValue=""
        self.InputButton=QPushButton("Input")
        self.OutPutButton=QPushButton("output")
        self.LoadButton=QPushButton("Load")
        self.SubmitButton=QPushButton("Submit")
        self.msg=QMessageBox()

        self.InputButton.setStyleSheet("border-radius:10px ;border:1px solid black;background-color: yellow")
        self.InputButton.clicked.connect(self.input_open_dialog)
        self.LoadButton.clicked.connect(self.load_file)
        self.OutPutButton.clicked.connect(self.output_open_dialog)
        self.SubmitButton.clicked.connect(self.submit_file)

        

        layout=QVBoxLayout()
        layout.addWidget(self.InputButton)
        layout.addWidget(self.OutPutButton)
        layout.addWidget(self.LoadButton)
        layout.addWidget(self.SubmitButton)
        

        container=QWidget()
        container.setLayout(layout)

        self.setFixedSize(QSize(200,200))
        
        self.setCentralWidget(container)

    
    def output_open_dialog(self):
        self.OutPut = QFileDialog.getOpenFileName(
            self,
            "Open File",
            "${HOME}",
            "All Files (*);; Excel file(*.xlsx)",
        )
        print(self.OutPut[0])

    def input_open_dialog(self):
        self.InPut = QFileDialog.getOpenFileName(
            self,
            "Open File",
            "${HOME}",
            "All Files (*);; Excel file(*.xlsx)",
        )
        print(self.InPut[0])
        

    def load_file(self):
        if self.InPut:
            self.InPutValue=user_input(self.InPut[0])
            self.msg.setWindowTitle("System")
            self.msg.setText(f"{self.InPutValue[1]}")
            self.msg.setIcon(QMessageBox.Icon.Information)
            self.msg.exec()
        else:
            self.msg.setWindowTitle("System")
            self.msg.setText(f"Please Enter your file")
            self.msg.setIcon(QMessageBox.Icon.Critical)
            self.msg.exec()
    
    def submit_file(self):
        if self.OutPut:
            output(process(self.InPutValue),self.OutPut[0])
            self.msg.setWindowTitle("System")
            self.msg.setText("Done")
            self.msg.setIcon(QMessageBox.Icon.Information)
            self.msg.exec()
        else:
            self.msg.setWindowTitle("System")
            self.msg.setText(f"Please Enter your file")
            self.msg.setIcon(QMessageBox.Icon.Critical)
            self.msg.exec()



app = QApplication(sys.argv) 
window = MainWindow()
window.show()

app.exec()
