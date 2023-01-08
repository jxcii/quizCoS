from typing import Dict
from PyQt5.QtCore import QSize, Qt
from PyQt5.QtGui import QFont, QKeySequence, QIcon
from PyQt5.QtWidgets import QHBoxLayout, QApplication, QMainWindow, QPushButton, QWidget, QVBoxLayout, QLabel, QLineEdit, QProgressBar, QShortcut
import sys
import glob
from openpyxl import load_workbook
import re
import random
import os
##########################################
progressCSS = """QProgressBar
{
border: solid grey;
border-radius: 15px;
color: black;
}
QProgressBar::chunk 
{
background-color: #05B8CC;
border-radius :15px;
} """
buttoncss = """
QLabel{
border: 3px solid #40E0D0;
border-radius: 10px;
background-color: #40E0D0;
font-weight: 658;
}
QLabel:hover{
background-color: #008080;
}
"""
buttoncssGreen = """
QLabel{
border: 3px solid SpringGreen;
border-radius: 10px;
background-color: SpringGreen;
font-weight: 658;
}
"""
buttoncssRed = """
QLabel{
border: 3px solid Tomato;
border-radius: 10px;
background-color: Tomato;
font-weight: 658;
}
"""
##########################################
def readFiles(a, b, c)->Dict:
    lib:Dict = {}
    ptr = 1
    for j in Files:
        i = a
        print(os.getcwd()+r"\\"+j)
        wb = load_workbook(os.getcwd()+r"\\"+j)
        sheet = wb['list1']
        while  sheet[f"A{i}"].value is not None:
            if c == 'v': 
                lib[f"{ptr}"] = { "question" : sheet[f"A{i}"].value, "a" : sheet[f"A{i+1}"].value,  "b" : sheet[f"A{i+2}"].value,  "c" : sheet[f"A{i+3}"].value,  "d" : sheet[f"A{i+4}"].value  }
            else:
                lib[f"{ptr}"] = { "question" : sheet[f"A{i}"].value, "a" : sheet[f"B{i}"].value,  "b" : sheet[f"C{i}"].value,  "c" : sheet[f"D{i}"].value,  "d" : sheet[f"E{i}"].value  }
            i += b
            ptr+=1
        wb.close()
    return lib
##########################################
Files = []
print(os.getcwd())
##########################################
class MyApp(QMainWindow):
    def __init__(self):
        self.libx:Dict = {}
        self.okno = 0
        super().__init__()
        #self.setWindowFlag(Qt.FramelessWindowHint)
        #self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowOpacity(0.9)
        self.setWindowTitle('quizCoS')
        self.setWindowIcon(QIcon('icon.png'))
        self.setStyleSheet("color: white; background-color: Teal; font-size: 18px;")
        self.setFont(QFont('Arial', 14))
        self.getUI()

    def getUI(self):
        if self.centralWidget() is not None:
            self.centralWidget().deleteLater()
        main = QWidget()
        self.config = QLineEdit('n-2 -v next-9')
        self.config.setPlaceholderText("config: ")

        vbox = QVBoxLayout()
        vbox.addLayout(self.core())
        vbox.addWidget(self.config, alignment=Qt.AlignCenter | Qt.AlignBottom)
        start = QPushButton("start")
        vbox.addWidget(start, alignment=Qt.AlignCenter | Qt.AlignBottom)
        start.clicked.connect(self.middleBuild)
        consolejx = QLabel("consolejx")
        consolejx.setStyleSheet('font-size: 10px;')
        vbox.addWidget(consolejx, alignment=Qt.AlignLeft | Qt.AlignBottom)

        main.setLayout(vbox)
        main.setContentsMargins(10, 10, 10, 10)

        self.setCentralWidget(main)
        self.setMinimumSize(300, 200)

    def core(self)->QHBoxLayout:
        hboX = QHBoxLayout()
        self.files = glob.glob(os.getcwd()+'/*.xlsx')
        for i in self.files:
            Qi = QWidget()
            name = re.search('[a-z]*.xlsx', i)
            vboX = QVBoxLayout()
            button : QPushButton = QPushButton("")
            button.setFixedSize(QSize(60, 70))
            button.setStyleSheet(" background-image: url('excel.png'); background-repeat: no-repeat; border-radius: 10px;")
            button.clicked.connect(self.checked_btn)
            label1 = QLabel(name.group(0))
            label1.setFixedSize(QSize(60, 20))
            label1.setStyleSheet('font-size: 12px;')
            label1.setAlignment(Qt.AlignCenter)
            vboX.addWidget(button)
            vboX.addWidget(label1)
            vboX.setAlignment(Qt.AlignCenter)
            Qi.setLayout(vboX)
            print(Qi.styleSheet())
            hboX.addWidget(Qi)
        return hboX

    def checked_btn(self):
        if self.sender().parentWidget().styleSheet() == "":
            self.sender().parentWidget().setStyleSheet("background-color: MediumAquamarine; border-radius: 20px;")
            Files.append(self.sender().parentWidget().children()[2].text())
        else:
            self.sender().parentWidget().setStyleSheet("")
            Files.remove(self.sender().parentWidget().children()[2].text())
        print(Files)
    
    def middleBuild(self):
        self.fileconfig = self.config.text()
        self.build()

    def build(self):
        self.okno = 1
        config = self.fileconfig
        a = re.search("(?<=n-)[0-9]*", config)
        a = int(a.group(0))
        b = re.search(r"v-*", config)
        if b is None:
            b = 'h'
        else: 
            b = 'v'
        c = re.search("(?<=next-)[0-9]*", config)
        c = int(c.group(0))

        self.libx.clear()
        self.libx = readFiles(a, c-a, b)

        self.centralWidget().deleteLater()
        main = QWidget()

        vbox = QVBoxLayout()

        Q1 = QWidget()
        slide = QHBoxLayout()
        self.a1 = QLineEdit()
        self.a2 = QLineEdit()
        self.a1.setContentsMargins(70, 0, 0, 0)
        self.a2.setContentsMargins(0, 0, 70, 0)
        self.a1.setMinimumHeight(30)
        self.a2.setMinimumHeight(30)
        self.a1.setAlignment(Qt.AlignCenter)
        self.a2.setAlignment(Qt.AlignCenter)
        slide.addWidget(self.a1)
        slide.addWidget(self.a2)
        Q1.setLayout(slide)

        self.prs = QProgressBar()
        self.prs.setAlignment(Qt.AlignCenter)
        self.prs.setStyleSheet(progressCSS)
        self.lbl1 = QLabel(f'{len(self.libx)} questions')
        self.lbl1.setWordWrap(True)
        self.lbl1.setAlignment(Qt.AlignCenter)

        self.btn1 = QLabel("a")
        self.btn2 = QLabel("b")
        self.btn3 = QLabel("c")
        self.btn4 = QLabel("d")
        self.btn1.setWordWrap(True)
        self.btn2.setWordWrap(True)
        self.btn3.setWordWrap(True)
        self.btn4.setWordWrap(True)
        self.btn1.setAlignment(Qt.AlignCenter)
        self.btn2.setAlignment(Qt.AlignCenter)
        self.btn3.setAlignment(Qt.AlignCenter)
        self.btn4.setAlignment(Qt.AlignCenter)
        self.btn1.setStyleSheet(buttoncss)
        self.btn2.setStyleSheet(buttoncss)
        self.btn3.setStyleSheet(buttoncss)
        self.btn4.setStyleSheet(buttoncss)

        self.btns = [self.btn1, self.btn2, self.btn3, self.btn4]

        self.index = QHBoxLayout()
        self.N = QPushButton("next")
        self.indexator = QLabel()
        self.index.addWidget(self.indexator, alignment=Qt.AlignLeft)
        self.index.addWidget(self.N, alignment=Qt.AlignRight)

        self.N.clicked.connect(self.MiddleWork)
        sh = QShortcut(QKeySequence("Return"), self)
        sh.activated.connect(self.MiddleWork)
        blick = QPushButton("Start")
        blick.setMaximumWidth(150)
        blick.clicked.connect(self.questionStart)
        info = QLabel('return is Esc')
        info.setStyleSheet('font-size: 14px')
        vbox.addWidget(info, alignment=Qt.AlignLeft)
        vbox.addWidget(self.lbl1) 
        vbox.addWidget(Q1)
        vbox.addWidget(blick, alignment=Qt.AlignCenter)
        vbox.addWidget(self.prs)
        vbox.addWidget(self.btn1) 
        vbox.addWidget(self.btn2) 
        vbox.addWidget(self.btn3) 
        vbox.addWidget(self.btn4)
        vbox.addLayout(self.index)
        
        main.setLayout(vbox)
        self.setCentralWidget(main)
        self.setMinimumSize(500, 400)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            if self.okno == 2:
                self.okno = 1
                self.build()
            elif self.okno == 1:
                self.okno = 0
                Files.clear()
                self.getUI()
            else:
                self.close()

    def questionStart(self):
        if self.a1.text()!="" and self.a2.text()!="":
            if int(self.a2.text())>len(self.libx):
                self.a2.setText("")
            elif int(self.a1.text())<1:
                self.a1.setText("")
            elif int(self.a1.text()) > int(self.a2.text()):
                self.a1.setText("")
                self.a2.setText("")
            else:
                self.start = int(self.a1.text())
                self.end = int(self.a2.text())
                self.prs.setMaximum(self.end-self.start + 1)
                self.countTest = self.end-self.start + 1
                self.prs.setFormat(f"0/{self.countTest}")
                self.prs.setValue(0)
                self.CountTrueAnswer = 0
                self.next()

    def next(self):
        self.shortcuts = [QShortcut(QKeySequence("1"), self), QShortcut(QKeySequence("2"), self), QShortcut(QKeySequence("3"), self), QShortcut(QKeySequence("4"), self) ]
        self.boolT = False
        self.lbl1.setText(self.libx[f"{self.start}"]['question'])
        mas = random.sample([ self.libx[f"{self.start}"]['a'], self.libx[f"{self.start}"]['b'], self.libx[f"{self.start}"]['c'],  self.libx[f"{self.start}"]['d']], 4)
        ptr = 0
        for i in mas:
            self.btns[ptr].setText(str(i))
            if i == self.libx[f"{self.start}"]['a']:
                self.btns[ptr].mousePressEvent = lambda e, x=self.btns[ptr]: self.trueAnswer(e,x)
                self.shortcuts[ptr].activated.connect(lambda e = None, x=self.btns[ptr]: self.trueAnswer(e,x))
            else:
                self.btns[ptr].mousePressEvent = lambda e, x=self.btns[ptr]: self.falseAnswer(e,x)
                self.shortcuts[ptr].activated.connect(lambda e = None, x=self.btns[ptr]: self.falseAnswer(e,x))
            ptr+=1

    def trueAnswer(self, e, btn):
        if self.boolT == False:
            self.CountTrueAnswer+=1
        btn.setStyleSheet(buttoncssGreen)
        self.boolT = True
    
    def falseAnswer(self, e, btn):
        btn.setStyleSheet(buttoncssRed)
        self.boolT = True
       
    def MiddleWork(self):
        self.indexator.setText(f'Result = {self.CountTrueAnswer}/{self.countTest}')
        for i in self.btns:
            i.setStyleSheet(buttoncss)
        self.start+=1
        self.prs.setValue(self.prs.value()+1)
        self.prs.setFormat(f"{self.countTest - (self.end-self.start + 1)}/{self.countTest}")
        if self.start == self.end + 1:
            self.endAP()
        else:
            for i in self.shortcuts:
                i.deleteLater()
            self.next()
    
    def endAP(self):
        self.okno = 2
        if self.centralWidget() is not None:
            self.centralWidget().deleteLater()
        main = QWidget()

        vbox = QVBoxLayout()
        Result = QLabel(f"{self.CountTrueAnswer}/{self.countTest} = {round(self.CountTrueAnswer/self.countTest,4)*100}%")
        Result.setStyleSheet("font-size: 26px;")
        Result.setAlignment(Qt.AlignCenter)
        vbox.addWidget(Result, alignment=Qt.AlignCenter)

        main.setLayout(vbox)
        main.setContentsMargins(10, 10, 10, 10)

        self.setCentralWidget(main)
        self.setMinimumSize(300, 200)
        



#####################################
app = QApplication(sys.argv)
w = MyApp()
w.show()
app.exec()
#####################################
