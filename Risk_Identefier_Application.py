from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMessageBox, QLineEdit, QProgressBar, QLabel, QFileDialog, QCheckBox, QMenuBar, QStatusBar
import re
import pandas as pd
import os
import time
from datetime import datetime
from functools import partial

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
        self.progressBar.setGeometry(QtCore.QRect(50, 835, 200, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.progressBar.hide()
        self.progressBarLabel = QLabel(self.centralwidget)
        self.progressBarLabel.setGeometry(QtCore.QRect(500, 835, 300, 25))        
        font = QtGui.QFont()
        font.setPointSize(14)        
        self.progressBarLabel.setFont(font)
        self.progressBarLabel.setObjectName("progressBarLabel")
        self.progressBarLabel.setStyleSheet("color: #effeff;")
        self.hideProgNumLabel = QLabel(self.centralwidget)
        self.hideProgNumLabel.setGeometry(QtCore.QRect(220, 835, 50, 23))
        self.hideProgNumLabel.setObjectName("hideProgNumLabel")
        self.hideProgNumLabel.setStyleSheet("background: #707070;")
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
        self.runButton.setGeometry(QtCore.QRect(282, 735, 561, 81))
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
        font.setPointSize(12)
        self.positiveInput.setFont(font)       
        self.positiveInput.setObjectName("positiveInput")
        self.positiveInput.setStyleSheet("background: gray; border-radius: 10px;")        
        self.negativeInput = QLineEdit(self.centralwidget)
        self.negativeInput.setEnabled(False)
        self.negativeInput.setGeometry(QtCore.QRect(500, 630, 331, 71))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.negativeInput.setFont(font)
        self.negativeInput.setObjectName("negativeInput")
        self.negativeInput.setStyleSheet("background: gray; border-radius: 10px")        
        self.sheetInput = QLineEdit(self.centralwidget)
        self.sheetInput.setGeometry(QtCore.QRect(30, 580, 171, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.sheetInput.setFont(font)
        self.sheetInput.setObjectName("sheetInput")
        self.sheetInput.setStyleSheet("background: white; border-radius: 10px")        
        self.columnInput = QLineEdit(self.centralwidget)
        self.columnInput.setGeometry(QtCore.QRect(250, 580, 171, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.columnInput.setFont(font)
        self.columnInput.setText("")
        self.columnInput.setObjectName("columnInput")
        self.columnInput.setStyleSheet("background: white; border-radius: 10px")        
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
        self.selectedFileLabel.setGeometry(QtCore.QRect(225, 470, 270, 25))
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
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Risks Identifier"))
        self.browseButton.setText(_translate("MainWindow", "Browse For File"))
        self.title.setText(_translate("MainWindow", "Risks Identifier"))
        self.runButton.setText(_translate("MainWindow", "Run Tool"))
        self.positiveCheckBox.setText(_translate("MainWindow", "Default Positive Words"))
        self.negativeCheckBox.setText(_translate("MainWindow", "Default Negative Words"))
        self.sheetNameLabel.setText(_translate("MainWindow", "Sheet Name:"))
        self.columnNameLabel.setText(_translate("MainWindow", "Column Name:"))
        self.progressBarLabel.setText(_translate("MainWindow", ""))
        self.positiveWordsLabel.setText(_translate("MainWindow", "Positive words separated by commas (words must exist in documents):"))
        self.negativeWordsLabel.setText(_translate("MainWindow", "Negative words separated by commas (words must exist in documents):"))


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        self.msg = QtWidgets.QMessageBox()
        self.Actionlistenr()

        thread = QtCore.QThread(self)
        thread.start()
        self.m_worker = Worker()
        self.m_worker.moveToThread(thread)
        self.alreadyPressed = False

    def Actionlistenr(self):
        self.browseButton.clicked.connect(self.browseAction)
        self.positiveCheckBox.stateChanged.connect(self.checkBoxAction)
        self.negativeCheckBox.stateChanged.connect(self.checkBoxAction)       
        self.runButton.clicked.connect(self.onButtonClick)
        self.positiveInput.returnPressed.connect(self.onButtonClick)
        self.negativeInput.returnPressed.connect(self.onButtonClick)
        self.sheetInput.returnPressed.connect(self.onButtonClick)
        self.columnInput.returnPressed.connect(self.onButtonClick)

    def browseAction(self):
        if not self.alreadyPressed:
            self.fileName, _ = QFileDialog.getOpenFileName(self.centralwidget,"QFileDialog.getOpenFileName()", "","Excel(*.xls *.xlsx)")
            url = QtCore.QUrl.fromLocalFile(self.fileName)
            if self.fileName:
                self.selectedFileLabel.setText("Selected File: "+url.fileName().lower())
        else:
            self.msg.setIcon(QMessageBox.Critical)
            self.msg.setText("Application is already running!")
            self.msg.setWindowTitle("Error")
            self.msg.exec()


    def checkBoxAction(self):
        if self.negativeCheckBox.isChecked():
            self.negativeInput.setEnabled(False)   
            self.negativeInput.setStyleSheet("background: gray; border-radius: 10px;")
        else:
            self.negativeInput.setEnabled(True)
            self.negativeInput.setStyleSheet("background: white; border-radius: 10px;")

        if self.positiveCheckBox.isChecked():
            self.positiveInput.setEnabled(False)  
            self.positiveInput.setStyleSheet("background: gray; border-radius: 10px;")
        else:
            self.positiveInput.setEnabled(True)
            self.positiveInput.setStyleSheet("background: white; border-radius: 10px;")


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
        return [[word for word in simple_preprocess(str(doc)) if word not in self.stop_words] for doc in texts]

    def request_information(self):
        if self.alreadyPressed:
            self.msg.setIcon(QMessageBox.Critical)
            self.msg.setText("Application is already running!")
            self.msg.setWindowTitle("Error")
            self.msg.exec()
        else:            
            if self.selectedFileLabel.text() is "":
                # return print('Please select a document first!')
                self.msg.setIcon(QMessageBox.Critical)
                self.msg.setText("Please select a document first!")
                self.msg.setWindowTitle("Error")
                return self.msg.exec()
            else:
                xl = pd.ExcelFile(self.fileName)
                if self.sheetInput.text() == "":
                    # return print('Please provide a sheet name!')
                    self.msg.setIcon(QMessageBox.Critical)
                    self.msg.setText("Please provide a sheet name!")
                    self.msg.setWindowTitle("Error")     
                    return self.msg.exec()       
                elif not self.column_in_columns(map(str.lower, xl.sheet_names), self.sheetInput.text().lower()):
                    # return print('The sheet name you provided does not exist!')
                    self.msg.setIcon(QMessageBox.Critical)
                    self.msg.setText("The sheet name you provided does not exist!")
                    self.msg.setWindowTitle("Error")   
                    return self.msg.exec()
                else:  
                    df = pd.read_excel(self.fileName)
                    df.columns = map(str.lower, df.columns)
                    # Convert the data-frame(df) to a list
                    if self.columnInput.text() == "":
                        # return print('Please provide a column name!')
                        self.msg.setIcon(QMessageBox.Critical)
                        self.msg.setText("Please provide a column name!")
                        self.msg.setWindowTitle("Error")
                        return self.msg.exec()
                    elif not self.column_in_columns(df.columns, self.columnInput.text().lower()):
                        # return print('The column name you provided does not exist!')
                        self.msg.setIcon(QMessageBox.Critical)
                        self.msg.setText("The column name you provided does not exist!")
                        self.msg.setWindowTitle("Error")
                        return self.msg.exec()
                    else:    
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
                                # return print('error')

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
                                # return print('error')
                            positive = ['injuries', 'fail', 'dangerous', 'oil', 'incident', 'struck',
                                        'hit', 'derail', 'incorrect', 'collision', 'fatal', 'leaking', 'tamper', 'emergency',
                                        'gas', 'collided', 'damage', 'risk', 'broken', 'break']
                        else:
                            positive_words = self.positiveInput.text().replace(" ", "").split(",")
                            positive = positive_words
                            negative_words = self.negativeInput.text().replace(" ", "").split(",")
                            negative = negative_words
                            if positive[0] is '' or negative[0] is '':
                                self.msg.setIcon(QMessageBox.Critical)
                                self.msg.setText("Positive/Negative words can't be empty, write something or use default values!")
                                self.msg.setWindowTitle("Error")
                                return self.msg.exec() 
                                # return print('error')
                        wrapper = partial(self.m_worker.task, self.fileName, self.columnInput, self.sent_to_words, self.remove_stopwords, self.array_length, df, positive, negative, self.progressBarLabel, self.progressBar, self.hideProgNumLabel)
                        QtCore.QTimer.singleShot(0, wrapper)
                        self.progressBar.show()  
                        self.progressBarLabel.setText("Training Model...")                  
                        self.calc = External()        
                        self.calc.countChanged.connect(self.onCountChanged)
                        self.calc.start()
                        self.alreadyPressed = True
                        # if wrapper.finished:                        
                        #     self.msg.setIcon(QMessageBox.Information)
                        #     self.msg.setText("Risk words have been successfully tagged. A new excel file named \""+os.path.splitext(os.path.basename(self.fileName))[0]+"RisksIdentified \""+" is generated to \""+os.path.dirname(os.path.realpath(__file__))+"\" with a new column 'Potential Risks'")
                        #     self.msg.setWindowTitle("Success!")
                        #     return self.msg.exec()                 

    
    @QtCore.pyqtSlot()
    def onButtonClick(self):        
        self.request_information()


    @QtCore.pyqtSlot(int)
    def onCountChanged(self, value):
        self.progressBar.setValue(value)


class Worker(QtCore.QObject):
    
    @QtCore.pyqtSlot(str)
    def task(self, fileName, columnInput, sent_to_words, remove_stopwords, array_length, df, positive, negative, progressBarLabel, progressBar, hideProgNumLabel):        

        # Initialize start time to run script
        start_time = datetime.now()
        
        data = df[columnInput.text().lower()].values.tolist()

        # Remove the new line character and single quotes
        data = [re.sub(r'\s+', ' ', str(sent)) for sent in data]
        data = [re.sub("\'", "", str(sent)) for sent in data]

        # Convert our data to a list of words. Now, data_words is a 2D array,
        # each index contains a list of words
        data_words = list(sent_to_words(data))

        # Remove the stop words
        data_words_nostops = remove_stopwords(data_words)
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

        similar_words_size = array_length(model.wv.most_similar(positive=positive, negative=negative, topn=0))
        # text_file = open("Output.txt", "w")

        progressBarLabel.setText("Finding risks...")
        for i in model.wv.most_similar(positive=positive, negative=negative, topn=similar_words_size):
            if len(i[0]) > 2:
                risks.append(i)

        highest = risks[0][1]

        progressBarLabel.setText("Tagging documents with risks...")
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

        progressBarLabel.setText("Removing duplicates...")
        for key,value in risksWithDuplicates.items():
            newValue = list(dict.fromkeys(risksWithDuplicates[key]))
            newValue = {key: newValue}
            risksWithDuplicates.update(newValue)
        
        
        progressBarLabel.setText("Creating new file...")
        base = os.path.basename(fileName)
        name = os.path.splitext(base)[0]
        extension = os.path.splitext(base)[1]               
        df['Potential Risks'] = df.index.map(risksWithDuplicates)

        cols = list(df.columns.values) 
        cols.pop(cols.index(columnInput.text()))
        cols.pop(cols.index('Potential Risks'))
        df = df[cols+[columnInput.text(), 'Potential Risks']]


        df.to_excel(name+"RisksIdentified"+extension.lower())
        # text_file.close()                   
        
        # msg.setIcon(QMessageBox.Information)
        # msg.setText("Risk words have been successfully tagged. A new excel file named \""+name+"RisksIdentified \""+" is generated to \""+os.path.dirname(os.path.realpath(__file__))+"\" with a new column 'Potential Risks'")
        # msg.setWindowTitle("Success!")              
        progressBar.hide()
        # progressBarLabel
        progressBarLabel.setGeometry(QtCore.QRect(20, 835, 1200, 25))
        hideProgNumLabel.hide()
        font = QtGui.QFont()
        font.setPointSize(10)
        progressBarLabel.setText("Risk words have been successfully tagged. A new excel file named \""+name+"RisksIdentified \""+" is generated to \""+os.path.dirname(os.path.realpath(__file__))+"\" with a new column 'Potential Risks'")
        progressBarLabel.setFont(font)
        # Print time to run script
        print("Time taken to run script: ", datetime.now() - start_time)
        # finished = True
        # return msg.exec()

class External(QThread):
    """
    Runs a counter thread.
    """
    countChanged = pyqtSignal(int)

    def run(self):
        count = 0
        while count < TIME_LIMIT:
            count +=20
            time.sleep(1)
            self.countChanged.emit(count)
            if count >= TIME_LIMIT:
                count=0



if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())