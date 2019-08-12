# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'test.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMessageBox, QLineEdit, QProgressBar, QLabel, QFileDialog, QCheckBox, QMenuBar, QStatusBar
import re
import pandas as pd
import os
import time
from datetime import datetime

# Gensim
import gensim
from gensim.utils import simple_preprocess

# Stop words
from nltk.corpus import stopwords

from collections import defaultdict

TIME_LIMIT = 100
class Ui_MainWindow(object):    

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(1125, 903)
        MainWindow.setAutoFillBackground(False)        
        self.fileName=""
        self.msg = QMessageBox()
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.centralwidget.setStyleSheet("background: #707070;")
        self.progressBar = QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(180, 690, 118, 23))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        self.browseButton = QPushButton(self.centralwidget)
        self.browseButton.setEnabled(True)
        self.browseButton.setGeometry(QtCore.QRect(30, 460, 171, 41))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.browseButton.setFont(font)
        self.browseButton.setAutoDefault(False)        
        self.browseButton.setFlat(False)
        self.browseButton.setObjectName("browseButton")
        self.browseButton.setStyleSheet("QPushButton{background: #914ed4; border-radius: 10px; color: #effeff; border: 3px outset black;} QPushButton:hover{background : #4838e8; border: 3px outset white;};")
        self.titlePicture = QLabel(self.centralwidget)
        self.titlePicture.setGeometry(QtCore.QRect(0, 0, 1154, 347))
        self.titlePicture.setText("")
        self.titlePicture.setStyleSheet("background-image: url('background.jpg')")
        self.titlePicture.setObjectName("titlePicture")
        self.title = QLabel(self.centralwidget)
        self.title.setGeometry(QtCore.QRect(372, 340, 381, 61))
        self.title.setStyleSheet("background: transparent; color: #effeff;")
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(40)
        font.setBold(True)
        font.setWeight(75)
        self.title.setFont(font)
        self.title.setObjectName("title")
        self.runButton = QPushButton(self.centralwidget)
        self.runButton.setGeometry(QtCore.QRect(282, 780, 561, 81))
        font = QtGui.QFont()
        font.setPointSize(20)
        self.runButton.setFont(font)
        self.runButton.setObjectName("runButton")
        self.runButton.setStyleSheet("QPushButton{background: #914ed4; border-radius: 10px; color: #effeff; border: 3px outset black;} QPushButton:hover{background : #4838e8; border: 3px outset white;};")
        self.runButton.setDefault(True)        
        self.positiveCheckBox = QCheckBox(self.centralwidget)
        self.positiveCheckBox.setGeometry(QtCore.QRect(860, 530, 191, 20))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.positiveCheckBox.setFont(font)
        self.positiveCheckBox.setChecked(True)
        self.positiveCheckBox.setObjectName("positiveCheckBox")
        self.positiveCheckBox.setStyleSheet("QCheckBox{color:#effeff}")
        self.negativeCheckBox = QCheckBox(self.centralwidget)
        self.negativeCheckBox.setEnabled(True)
        self.negativeCheckBox.setGeometry(QtCore.QRect(860, 650, 191, 20))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.negativeCheckBox.setFont(font)
        self.negativeCheckBox.setAutoFillBackground(False)
        self.negativeCheckBox.setChecked(True)
        self.negativeCheckBox.setObjectName("negativeCheckBox")
        self.negativeCheckBox.setStyleSheet("color:#effeff")
        self.positiveInput = QLineEdit(self.centralwidget)
        self.positiveInput.setEnabled(False)
        self.positiveInput.setGeometry(QtCore.QRect(500, 510, 331, 71))
        font = QtGui.QFont()
        font.setPointSize(70)
        self.positiveInput.setFont(font)       
        self.positiveInput.setObjectName("positiveInput")
        self.positiveInput.setText("XXXXXX")
        self.positiveInput.setStyleSheet("background: white; border-radius: 10px;")
        self.positiveInput.returnPressed.connect(self.run)
        self.negativeInput = QLineEdit(self.centralwidget)
        self.negativeInput.setEnabled(False)
        self.negativeInput.setGeometry(QtCore.QRect(500, 630, 331, 71))
        font = QtGui.QFont()
        font.setPointSize(70)
        self.negativeInput.setFont(font)
        self.negativeInput.setObjectName("negativeInput")
        self.negativeInput.setText("XXXXXX")
        self.negativeInput.setStyleSheet("background: white; border-radius: 10px")
        self.negativeInput.returnPressed.connect(self.run)
        self.sheetInput = QLineEdit(self.centralwidget)
        self.sheetInput.setGeometry(QtCore.QRect(30, 580, 171, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.sheetInput.setFont(font)
        self.sheetInput.setObjectName("sheetInput")
        self.sheetInput.setStyleSheet("background: white; border-radius: 10px")
        self.sheetInput.returnPressed.connect(self.run)
        self.columnInput = QLineEdit(self.centralwidget)
        self.columnInput.setGeometry(QtCore.QRect(250, 580, 171, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.columnInput.setFont(font)
        self.columnInput.setText("")
        self.columnInput.setObjectName("columnInput")
        self.columnInput.setStyleSheet("background: white; border-radius: 10px")
        self.columnInput.returnPressed.connect(self.run)
        self.sheetNameLabel = QLabel(self.centralwidget)
        self.sheetNameLabel.setGeometry(QtCore.QRect(30, 550, 131, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.sheetNameLabel.setFont(font)
        self.sheetNameLabel.setObjectName("sheetNameLabel")
        self.sheetNameLabel.setStyleSheet("color: #effeff;")
        self.columnNameLabel = QLabel(self.centralwidget)
        self.columnNameLabel.setGeometry(QtCore.QRect(250, 550, 131, 16))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.columnNameLabel.setFont(font)
        self.columnNameLabel.setObjectName("columnNameLabel")
        self.columnNameLabel.setStyleSheet("color: #effeff;")
        self.positiveWordsLabel = QLabel(self.centralwidget)
        self.positiveWordsLabel.setGeometry(QtCore.QRect(500, 480, 601, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.positiveWordsLabel.setFont(font)
        self.positiveWordsLabel.setObjectName("positiveWordsLabel")
        self.positiveWordsLabel.setStyleSheet("color: #effeff;")
        self.negativeWordsLabel = QLabel(self.centralwidget)
        self.negativeWordsLabel.setGeometry(QtCore.QRect(500, 600, 611, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.negativeWordsLabel.setFont(font)
        self.negativeWordsLabel.setObjectName("negativeWordsLabel")
        self.negativeWordsLabel.setStyleSheet("color: #effeff;")
        self.selectedFileLabel = QLabel(self.centralwidget)
        self.selectedFileLabel.setGeometry(QtCore.QRect(250, 470, 101, 16))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.selectedFileLabel.setFont(font)
        self.selectedFileLabel.setText("")
        self.selectedFileLabel.setObjectName("selectedFileLabel")
        self.selectedFileLabel.setStyleSheet("color: #effeff;")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1125, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.stop_words = stopwords.words('english')
        extended_stop_words = ['from', 're', 'use', 'any', 'also', 'known']
        self.stop_words.extend(extended_stop_words)
        self.retranslateUi(MainWindow)
        self.actionListener()
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


    # def onButtonClick(self):
    #     self.calc = External()        
    #     self.calc.countChanged.connect(self.onCountChanged)
    #     self.calc.start()
    #     self.calc2 = External2(self.selectedFileLabel, self.msg, self.sheetInput, self.columnInput, self.positiveCheckBox, self.negativeCheckBox, self.fileName, self.positiveInput, self.negativeInput)
    #     # self.calc2.window.connect(self.MainWindow)   
    #     # self.calc2.centralwidget.connect(self.centralwidget(MainWindow))        
    #     # self.calc2.selectedFileLabel.connect(self.selectedFileLabel)
    #     # self.calc2.msg.connect(self.msg)
    #     # self.calc2.sheetInput.connect(self.sheetInput)
    #     # self.calc2.columnInput.connect(self.columnInput)
    #     # self.calc2.positiveCheckbox.connect(self.positiveCheckbox)
    #     # self.calc2.negativeCheckbox.connect(self.negativeCheckbox)
    #     # self.calc2.fileName.connect(self.fileName)
    #     self.calc2.start()


    # def onCountChanged(self, value):
    #     self.progressBar.setValue(value)


    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.browseButton.setText(_translate("MainWindow", "Browse For File"))
        self.title.setText(_translate("MainWindow", "Risks Identifier"))
        self.runButton.setText(_translate("MainWindow", "Run Tool"))
        self.positiveCheckBox.setText(_translate("MainWindow", "Default Positive Words"))
        self.negativeCheckBox.setText(_translate("MainWindow", "Default Negative Words"))
        self.sheetNameLabel.setText(_translate("MainWindow", "Sheet Name:"))
        self.columnNameLabel.setText(_translate("MainWindow", "Column Name:"))
        self.positiveWordsLabel.setText(_translate("MainWindow", "Positive words separated by commas (words must exist in documents):"))
        self.negativeWordsLabel.setText(_translate("MainWindow", "Negative words separated by commas (words must exist in documents):"))


    def actionListener(self):
        self.browseButton.clicked.connect(self.browseAction)
        self.positiveCheckBox.stateChanged.connect(self.checkBoxAction)
        self.negativeCheckBox.stateChanged.connect(self.checkBoxAction)       
        self.runButton.clicked.connect(self.run)
        

    def browseAction(self):
        self.fileName, _ = QFileDialog.getOpenFileName(self.centralwidget,"QFileDialog.getOpenFileName()", "","Excel(*.xls *.xlsx)")
        url = QtCore.QUrl.fromLocalFile(self.fileName)
        if self.fileName:
            self.selectedFileLabel.setText(url.fileName().lower())  


    def checkBoxAction(self):
        if self.negativeCheckBox.isChecked():
            self.negativeInput.setEnabled(False)
            font = QtGui.QFont()
            font.setPointSize(70)
            self.negativeInput.setFont(font)            
            self.negativeInput.setText("XXXXXX");
        else:
            self.negativeInput.setEnabled(True)
            font = QtGui.QFont()
            font.setPointSize(12)
            self.negativeInput.setFont(font) 
            self.negativeInput.setText("");

        if self.positiveCheckBox.isChecked():
            self.positiveInput.setEnabled(False)
            font = QtGui.QFont()
            font.setPointSize(70)
            self.positiveInput.setFont(font)            
            self.positiveInput.setText("XXXXXX");
        else:
            self.positiveInput.setEnabled(True)
            font = QtGui.QFont()
            font.setPointSize(12)
            self.positiveInput.setFont(font) 
            self.positiveInput.setText("");


    def array_length(self, array):
        count = 0
        for i in array:
            count += 1
        return count


    def column_in_columns(self, column_names, text):
        for i in column_names:
            if text == i:
                return True
        return False


    def sent_to_words(self, sentences):
        # Tokenize each sentence into a list of words and remove unwanted characters
        for sentence in sentences:
            yield(gensim.utils.simple_preprocess(str(sentence), deacc=True))  


    def remove_stopwords(self, texts):
        # Removes stopwords in a text
        return [[word for word in simple_preprocess(str(doc)) if word not in stop_words] for doc in texts]


    def run(self):
        print(self.fileName)
        if self.selectedFileLabel.text() is "":
                self.msg.setIcon(QMessageBox.Critical)
                self.msg.setText("Please select a document first!")
                self.msg.setWindowTitle("Error")
                return self.msg.exec()
        else:
            xl = pd.ExcelFile(fileName)
            if self.sheetInput.text() == "":
                self.msg.setIcon(QMessageBox.Critical)
                self.msg.setText("Please provide a sheet name!")
                self.msg.setWindowTitle("Error")     
                return self.msg.exec()       
            elif not self.column_in_columns(map(str.lower, xl.sheet_names), self.sheetInput.text().lower()):
                self.msg.setIcon(QMessageBox.Critical)
                self.msg.setText("The sheet name you provided does not exist!")
                self.msg.setWindowTitle("Error")   
                return self.msg.exec()
            else:
                df = pd.read_excel(self.fileName)
                df.columns = map(str.lower, df.columns)

                # Convert the data-frame(df) to a list
                if self.columnInput.text() == "":
                    self.msg.setIcon(QMessageBox.Critical)
                    self.msg.setText("Please provide a column name!")
                    self.msg.setWindowTitle("Error")
                    return self.msg.exec()
                elif not self.column_in_columns(df.columns, self.columnInput.text().lower()):
                    self.msg.setIcon(QMessageBox.Critical)
                    self.msg.setText("The column name you provided does not exist!")
                    self.msg.setWindowTitle("Error")
                    return self.msg.exec()
                else:    
                    # Initialize start time to run script
                    start_time = datetime.now()

                    data = df[self.columnInput.text().lower()].values.tolist()

                    # Remove the new line character and single quotes
                    data = [re.sub(r'\s+', ' ', str(sent)) for sent in data]
                    data = [re.sub("\'", "", str(sent)) for sent in data]

                    # Convert our data to a list of words. Now, data_words is a 2D array,
                    # each index contains a list of words
                    data_words = list(self.sent_to_words(data))

                    # Remove the stop words
                    data_words_nostops = self.remove_stopwords(data_words)
                    # for index, i in enumerate(data_words_nostops):
                    #     for w in i:
                    #         if w == 'inj':
                    #             print(w, index)
                    # exit(0)
                    model = gensim.models.Word2Vec(
                            data_words_nostops,
                            alpha=0.1,
                            min_alpha=0.001,
                            size=250,
                            window=1,
                            min_count=2,
                            workers=10)
                    model.train(data_words_nostops, total_examples=len(data_words_nostops), epochs=10)

                    risks = []

                    if self.positiveCheckBox.isChecked() and self.negativeCheckBox.isChecked():
                        positive = ['injuries', 'fail', 'dangerous', 'oil', 'incident', 'struck',
                                    'hit', 'derail', 'incorrect', 'collision', 'fatal', 'leaking', 'tamper', 'emergency',
                                    'gas', 'collided', 'damage', 'risk', 'broken', 'break']

                        negative = ['train', 'westward', 'goods', 'calgary', 'car', 'automobile', 'appliance', 'south',
                                    'mileage', 'empl', 'crossing', 'local', 'track', 'signal', 'approximately', 'released', 'rail']

                    elif not self.positiveCheckBox.isChecked() and self.negativeCheckBox.isChecked():
                        positive_words = self.positiveInput.text().replace(" ", "").split(",")
                        positive = positive_words
                        if positive[0] is '':
                            self.msg.setIcon(QMessageBox.Critical)
                            self.msg.setText("Positive/Negative words can't be empty, write something or use default values!")
                            self.msg.setWindowTitle("Error")
                            return self.msg.exec()

                        negative = ['train', 'westward', 'goods', 'calgary', 'car', 'automobile', 'appliance', 'south',
                                    'mileage', 'empl', 'crossing', 'local', 'track', 'signal', 'approximately', 'released', 'rail']

                    elif self.positiveCheckBox.isChecked() and not self.negativeCheckBox.isChecked():
                        negative_words = self.negativeInput.text().replace(" ", "").split(",")
                        negative = negative_words
                        if negative[0] is '':
                            self.msg.setIcon(QMessageBox.Critical)
                            self.msg.setText("Positive/Negative words can't be empty, write something or use default values!")
                            self.msg.setWindowTitle("Error")
                            return self.msg.exec()
                        positive = ['injuries', 'fail', 'dangerous', 'oil', 'incident', 'struck',
                                    'hit', 'derail', 'incorrect', 'collision', 'fatal', 'leaking', 'tamper', 'emergency',
                                    'gas', 'collided', 'damage', 'risk', 'broken', 'break']
                    else:
                        positive_words = self.positiveInput.text().replace(" ", "").split(",")
                        positive = positive_words
                        negative_words = self.negativeInput.text().replace(" ", "").split(",")
                        negative = negative_words
                        if positive[0] is '' or negative[0] is '':
                            msg.setIcon(QMessageBox.Critical)
                            msg.setText("Positive/Negative words can't be empty, write something or use default values!")
                            msg.setWindowTitle("Error")
                            return msg.exec()                        
                        

                    similar_words_size = self.array_length(model.wv.most_similar(positive=positive, negative=negative, topn=0))
                    # text_file = open("Output.txt", "w")

                    for i in model.wv.most_similar(positive=positive, negative=negative, topn=similar_words_size):
                        if len(i[0]) > 2:
                            risks.append(i)

                    highest = risks[0][1]

                    risksWithDuplicates = defaultdict(list)
                    for document_number, i in enumerate(data_words_nostops):
                        for j in i:
                            if len(j) > 2:
                                for k in risks:
                                    if k[1] > highest/2 and k[0] in j and k[0][:3] == j[:3]:              
                                        risksWithDuplicates[document_number].append(j)                          
                                        # text_file.write("Document %i has the words %s from %s\n" % (document_number+1, j, k))
                                        break

                    risksWithDuplicates = dict(risksWithDuplicates)

                    for key,value in risksWithDuplicates.items():
                        newValue = list(dict.fromkeys(risksWithDuplicates[key]))
                        newValue = {key: newValue}
                        risksWithDuplicates.update(newValue)
                    
                    
                    base = os.path.basename(self.fileName)
                    name = os.path.splitext(base)[0]
                    extension = os.path.splitext(base)[1]               
                    df['Potential Risks'] = df.index.map(risksWithDuplicates)

                    cols = list(df.columns.values) 
                    cols.pop(cols.index(self.columnInput.text()))
                    cols.pop(cols.index('Potential Risks'))
                    df = df[cols+[self.columnInput.text(), 'Potential Risks']]


                    df.to_excel(name+"RisksIdentified"+extension.lower())
                    # text_file.close()                   
                    
                    self.msg.setIcon(QMessageBox.Information)
                    self.msg.setText("Risk words have been successfully tagged. A new excel file named \""+name+"RisksIdentified \""+" is generated to \""+os.path.dirname(os.path.realpath(__file__))+"\" with a new column 'Potential Risks'")
                    self.msg.setWindowTitle("Success!")

                    # Print time to run script
                    print("Time taken to run script: ", datetime.now() - start_time)
                    return self.msg.exec()



# class External(QThread):
#     """
#     Runs a counter thread.
#     """
#     countChanged = pyqtSignal(int)

#     def run(self):
#         count = 0
#         while count < TIME_LIMIT:
#             count +=1
#             time.sleep(1)
#             self.countChanged.emit(count)

# class External2(QThread):    
#     """
#     Runs a counter thread.
#     """
#     def __init__(self, selectedFileLabel, msg, sheetInput, columnInput, positiveCheckBox, negativeCheckBox, fileName, positiveInput, negativeInput):
#         super().__init__()
#         self.selectedFileLabel = selectedFileLabel
#         self.msg = msg
#         self.sheetInput = sheetInput
#         self.columnInput = columnInput
#         self.positiveCheckBox = positiveCheckBox
#         self.negativeCheckBox = negativeCheckBox
#         self.fileName = fileName
#         self.positiveInput = positiveInput
#         self.negativeInput = negativeInput

    

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

