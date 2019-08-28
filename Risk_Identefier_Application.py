import re
import pandas as pd
from shutil import move
import os
import time
from datetime import datetime
from functools import partial
from collections import defaultdict

# PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMessageBox, QLineEdit, QProgressBar, QLabel, QFileDialog, QCheckBox, QMenuBar, QStatusBar, QTextEdit, QRadioButton

# Gensim
import gensim
from gensim.utils import simple_preprocess

# Stop words
from nltk.corpus import stopwords

QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)


positiveList = ['injuries', 'fail', 'dangerous', 'oil', 'incident', 'struck',
                'hit', 'derail', 'incorrect', 'collision', 'fatal', 'leaking', 'tamper', 'emergency',
                'gas', 'collided', 'damage', 'risk', 'broken']

negativeList = ['train', 'westward', 'goods', 'however', 'calgary', 'car', 'automobile', 'appliance', 'south',
                'mileage', 'empl', 'crossing', 'local', 'track', 'signal', 'bell','approximately', 'released', 'rail', 'shop']
TIME_LIMIT = 100
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        screen = app.primaryScreen()
        screenHeight = screen.size().height()
        screenWidth = screen.size().width()
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(1125, 800)
        MainWindow.setAutoFillBackground(False)
        self.filedialog = QFileDialog()
        self.filedialog.setFixedSize(500,500)
        self.fileName=""
        self.msg = QMessageBox()        
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.centralwidget.setStyleSheet("background: #707070;")
        self.progressBar = QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(50, 835-103, 200, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.progressBar.hide()
        self.progressBarLabel = QLabel(self.centralwidget)
        self.progressBarLabel.setWordWrap(True);
        self.progressBarLabel.setGeometry(QtCore.QRect(500, 835-103, 300, 25))        
        font = QtGui.QFont()
        font.setPointSize(14)        
        self.progressBarLabel.setFont(font)
        self.progressBarLabel.setObjectName("progressBarLabel")
        self.progressBarLabel.setStyleSheet("color: #effeff;")
        self.hideProgNumLabel = QLabel(self.centralwidget)
        self.hideProgNumLabel.setGeometry(QtCore.QRect(215, 835-103, 50, 23))
        self.hideProgNumLabel.setObjectName("hideProgNumLabel")
        self.hideProgNumLabel.setStyleSheet("background: #707070;")
        self.radioButtonML = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButtonML.setGeometry(QtCore.QRect(30, 640-103, 180, 23))
        self.radioButtonML.setObjectName("radioButtonML")
        self.radioButtonML.setFont(font)
        self.radioButtonML.setStyleSheet("QRadioButton{color:#effeff} QRadioButton:indicator { image: url('images/uncheckedRadio.png'); width:20px; height: 150px;} QRadioButton:indicator:checked {image: url('images/radioChecked.png');}")
        self.radioButtonML.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.radioButtonML.setChecked(True)
        self.radioButtonSearch = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButtonSearch.setGeometry(QtCore.QRect(30, 700-103, 171, 23))
        self.radioButtonSearch.setObjectName("radioButtonSearch")
        self.radioButtonSearch.setFont(font)
        self.radioButtonSearch.setStyleSheet("QRadioButton{color:#effeff} QRadioButton:indicator { image: url('images/uncheckedRadio.png'); width:20px; height: 150px;} QRadioButton:indicator:checked {image: url('images/radioChecked.png');}")
        self.radioButtonSearch.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.browseButton = QPushButton(self.centralwidget)
        self.browseButton.setGeometry(QtCore.QRect(30, 460-103, 171, 41))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.browseButton.setFont(font)
        self.browseButton.setAutoDefault(False)        
        self.browseButton.setFlat(False)
        self.browseButton.setObjectName("browseButton")
        self.browseButton.setStyleSheet("QPushButton{background: #914ed4; border-radius: 10px; color: #effeff; border: 3px outset black;} QPushButton:hover{background : #4838e8; border: 3px outset white;};")
        self.browseButton.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.titlePicture = QLabel(self.centralwidget)
        self.titlePicture.setGeometry(QtCore.QRect(0, 0, 1154, 347))
        self.titlePicture.setText("")
        self.titlePicture.setStyleSheet("background-image: url('images/background.jpg')")
        self.titlePicture.setObjectName("titlePicture")
        self.title = QLabel(self.centralwidget)
        self.title.setGeometry(QtCore.QRect(550, 340-183, 381, 61))
        self.title.setStyleSheet("background: transparent; color: #effeff;")
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(40)
        font.setBold(True)
        font.setWeight(75)
        self.title.setFont(font)
        self.title.setObjectName("title")
        self.runButton = QPushButton(self.centralwidget)
        self.runButton.setGeometry(QtCore.QRect(282, 735-103, 561, 81))
        font = QtGui.QFont()
        font.setPointSize(20)
        self.runButton.setFont(font)
        self.runButton.setObjectName("runButton")
        self.runButton.setStyleSheet("QPushButton{background: #914ed4; border-radius: 10px; color: #effeff; border: 3px outset black;} QPushButton:hover{background : #4838e8; border: 3px outset white;};")
        self.runButton.setDefault(True)
        self.runButton.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.positiveCheckBox = QCheckBox(self.centralwidget)
        self.positiveCheckBox.setGeometry(QtCore.QRect(860, 530-103, 191, 20))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.positiveCheckBox.setFont(font)
        self.positiveCheckBox.setChecked(True)
        self.positiveCheckBox.setObjectName("positiveCheckBox")
        self.positiveCheckBox.setStyleSheet("QCheckBox{color:#effeff} QCheckBox:indicator { image: url('images/unchecked.png'); width:20px; height: 150px;} QCheckBox:indicator:checked {image: url('images/checked.png');}")
        self.positiveCheckBox.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.negativeCheckBox = QCheckBox(self.centralwidget)
        self.negativeCheckBox.setChecked(True)
        self.negativeCheckBox.setGeometry(QtCore.QRect(860, 650-103, 191, 20))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.negativeCheckBox.setFont(font)
        self.negativeCheckBox.setAutoFillBackground(False)
        self.negativeCheckBox.setChecked(True)
        self.negativeCheckBox.setObjectName("negativeCheckBox")
        self.negativeCheckBox.setStyleSheet("QCheckBox{color:#effeff} QCheckBox:indicator { image: url('images/unchecked.png'); width:20px; height: 150px;} QCheckBox:indicator:checked {image: url('images/checked.png');}")
        self.negativeCheckBox.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.positiveInput = QTextEdit(self.centralwidget)
        self.positiveInput.setGeometry(QtCore.QRect(500, 510-103, 331, 71))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.positiveInput.setFont(font)       
        self.positiveInput.setObjectName("positiveInput")
        self.positiveInput.setStyleSheet("background: white; border-radius: 10px;")
        self.positiveInput.setText(",".join(positiveList))
        self.negativeInput = QTextEdit(self.centralwidget)
        self.negativeInput.setGeometry(QtCore.QRect(500, 630-103, 331, 71))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.negativeInput.setFont(font)
        self.negativeInput.setObjectName("negativeInput")
        self.negativeInput.setStyleSheet("background: white; border-radius: 10px")
        self.negativeInput.setText(",".join(negativeList))      
        self.sheetInput = QLineEdit(self.centralwidget)
        self.sheetInput.setGeometry(QtCore.QRect(30, 580-103, 171, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.sheetInput.setFont(font)
        self.sheetInput.setObjectName("sheetInput")
        self.sheetInput.setStyleSheet("background: white; border-radius: 10px")        
        self.columnInput = QLineEdit(self.centralwidget)
        self.columnInput.setGeometry(QtCore.QRect(250, 580-103, 171, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.columnInput.setFont(font)
        self.columnInput.setText("")
        self.columnInput.setObjectName("columnInput")
        self.columnInput.setStyleSheet("background: white; border-radius: 10px")        
        self.sheetNameLabel = QLabel(self.centralwidget)
        self.sheetNameLabel.setGeometry(QtCore.QRect(30, 550-103, 131, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.sheetNameLabel.setFont(font)
        self.sheetNameLabel.setObjectName("sheetNameLabel")
        self.sheetNameLabel.setStyleSheet("color: #effeff;")
        self.columnNameLabel = QLabel(self.centralwidget)
        self.columnNameLabel.setGeometry(QtCore.QRect(250, 550-103, 131, 16))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.columnNameLabel.setFont(font)
        self.columnNameLabel.setObjectName("columnNameLabel")
        self.columnNameLabel.setStyleSheet("color: #effeff;")
        self.positiveWordsLabel = QLabel(self.centralwidget)
        self.positiveWordsLabel.setGeometry(QtCore.QRect(500, 480-103, 601, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.positiveWordsLabel.setFont(font)
        self.positiveWordsLabel.setObjectName("positiveWordsLabel")
        self.positiveWordsLabel.setStyleSheet("color: #effeff;")
        self.negativeWordsLabel = QLabel(self.centralwidget)
        self.negativeWordsLabel.setGeometry(QtCore.QRect(500, 600-103, 611, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.negativeWordsLabel.setFont(font)
        self.negativeWordsLabel.setObjectName("negativeWordsLabel")
        self.negativeWordsLabel.setStyleSheet("color: #effeff;")
        self.selectedFileLabel = QLabel(self.centralwidget)
        self.selectedFileLabel.setGeometry(QtCore.QRect(225, 470-103, 270, 25))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.selectedFileLabel.setFont(font)
        self.selectedFileLabel.setText("")
        self.selectedFileLabel.setObjectName("selectedFileLabel")
        self.selectedFileLabel.setStyleSheet("color: #effeff;")
        self.searchInput = QTextEdit(self.centralwidget)
        self.searchInput.setGeometry(QtCore.QRect(500, 580-103, 601, 100))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.searchInput.setFont(font)
        self.searchInput.setObjectName("sheetInput")
        self.searchInput.setStyleSheet("background: white; border-radius: 10px")
        self.searchInputLabel = QLabel(self.centralwidget)
        self.searchInputLabel.setGeometry(QtCore.QRect(500, 550-103, 611, 21))  
        font = QtGui.QFont()
        font.setPointSize(14)
        self.searchInputLabel.setFont(font)
        self.searchInputLabel.setObjectName("searchInputLabel")
        self.searchInputLabel.setStyleSheet("color: #effeff;")
        self.searchInput.hide()
        self.searchInputLabel.hide()

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
        self.radioButtonML.setText(_translate("MainWindow", "Machine Learning"))
        self.radioButtonSearch.setText(_translate("MainWindow", "Normal Search"))
        self.title.setText(_translate("MainWindow", "Risks Identifier"))
        self.runButton.setText(_translate("MainWindow", "Run Tool"))
        self.positiveCheckBox.setText(_translate("MainWindow", "Default Positive Words"))
        self.negativeCheckBox.setText(_translate("MainWindow", "Default Negative Words"))
        self.sheetNameLabel.setText(_translate("MainWindow", "Sheet Name:"))
        self.columnNameLabel.setText(_translate("MainWindow", "Column Name:"))
        self.progressBarLabel.setText(_translate("MainWindow", ""))
        self.positiveWordsLabel.setText(_translate("MainWindow", "Positive words separated by commas (words must exist in documents):"))
        self.negativeWordsLabel.setText(_translate("MainWindow", "Negative words separated by commas (words must exist in documents):"))
        self.searchInputLabel.setText(_translate("MainWindow", "Words to search for seperated by commas:"))

class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        self.msg = QtWidgets.QMessageBox()
        self.Actionlistenr()
        self.oldFileName = 'empty'

        thread = QtCore.QThread(self)
        thread.start()
        self.m_worker = Worker()
        self.m_worker.moveToThread(thread)
        self.alreadyPressed = [False]


    def Actionlistenr(self):
        self.browseButton.clicked.connect(self.browseAction)
        self.positiveCheckBox.stateChanged.connect(self.checkBoxAction)
        self.negativeCheckBox.stateChanged.connect(self.checkBoxAction)       
        self.runButton.clicked.connect(self.onButtonClick)
        self.sheetInput.returnPressed.connect(self.onButtonClick)
        self.columnInput.returnPressed.connect(self.onButtonClick)
        self.radioButtonML.toggled.connect(self.radioButtonAction)
        self.radioButtonSearch.toggled.connect(self.radioButtonAction)


    def radioButtonAction(self):
        if self.radioButtonML.isChecked():
            self.searchInput.hide()
            self.searchInputLabel.hide()
            self.positiveInput.show()
            self.positiveCheckBox.show()
            self.positiveWordsLabel.show()
            self.negativeInput.show()
            self.negativeCheckBox.show()
            self.negativeWordsLabel.show()
        else:
            self.searchInput.show()
            self.searchInputLabel.show()
            self.positiveInput.hide()
            self.positiveCheckBox.hide()
            self.positiveWordsLabel.hide()
            self.negativeInput.hide()
            self.negativeCheckBox.hide()
            self.negativeWordsLabel.hide()
            


    def browseAction(self):
        if not self.alreadyPressed[0]:
            self.fileName, _ = self.filedialog.getOpenFileName(self.centralwidget,"Open File", "","Excel(*.xls *.xlsx)")
            url = QtCore.QUrl.fromLocalFile(self.fileName)
            
            if self.fileName != '':
                self.selectedFileLabel.setText("Selected File: "+url.fileName().lower())
                self.oldFileName = self.fileName
            else:
                self.fileName = self.oldFileName
        else:
            self.msg.setIcon(QMessageBox.Critical)
            self.msg.setText("Application is already running!")
            self.msg.setWindowTitle("Error")
            self.msg.exec()


    def checkBoxAction(self):
        if self.negativeCheckBox.isChecked():
            self.negativeInput.setText(",".join(negativeList))
        else:
            self.negativeInput.setText("")

        if self.positiveCheckBox.isChecked():
            self.positiveInput.setText(",".join(positiveList))
        else:
            self.positiveInput.setText("")



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
        if self.alreadyPressed[0]:
            self.msg.setIcon(QMessageBox.Critical)
            self.msg.setText("Application is already running!")
            self.msg.setWindowTitle("Error")
            self.msg.exec()
        else:            
            if self.selectedFileLabel.text() is "":
                self.msg.setIcon(QMessageBox.Critical)
                self.msg.setText("Please select a document first!")
                self.msg.setWindowTitle("Error")
                return self.msg.exec()
            else:
                xl = pd.ExcelFile(self.fileName)
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
                        if self.radioButtonML.isChecked(): 
                            positive = self.positiveInput.toPlainText().replace(" ", "").split(",")
                            negative = self.negativeInput.toPlainText().replace(" ", "").split(",")
                            if positive[0] is '' or negative[0] is '':
                                self.msg.setIcon(QMessageBox.Critical)
                                self.msg.setText("Positive/Negative words can't be empty, write something or use default values!")
                                self.msg.setWindowTitle("Error")
                                return self.msg.exec()                            
                            wrapper = partial(self.m_worker.taskML, self.fileName, self.columnInput, self.sent_to_words, self.remove_stopwords, self.array_length, df, positive, negative, self.progressBarLabel, self.progressBar, self.hideProgNumLabel, self.runButton, self.alreadyPressed)

                        else:
                            if self.searchInput.toPlainText() == "":
                                self.msg.setIcon(QMessageBox.Critical)
                                self.msg.setText("Search words field cannot be empty!")
                                self.msg.setWindowTitle("Error")
                                return self.msg.exec()
                            else:                                       
                                wrapper = partial(self.m_worker.taskSearch, self.fileName, self.columnInput, self.sent_to_words, self.array_length, df, self.progressBarLabel, self.progressBar, self.hideProgNumLabel, self.runButton, self.searchInput, self.alreadyPressed)
                        
                        self.runButton.setGeometry(QtCore.QRect(282, 735-103, 561, 81))
                        self.progressBarLabel.setGeometry(QtCore.QRect(500, 835-103, 300, 25))        
                        font = QtGui.QFont()
                        font.setPointSize(14)        
                        self.progressBarLabel.setFont(font)
                        self.hideProgNumLabel.show()

                        QtCore.QTimer.singleShot(0, wrapper)
                        self.calc = External()        
                        self.calc.countChanged.connect(self.onCountChanged)
                        self.calc.start() 

                        self.progressBar.show()
                        self.alreadyPressed[0] = True     

    
    @QtCore.pyqtSlot()
    def onButtonClick(self):        
        self.request_information()


    @QtCore.pyqtSlot(int)
    def onCountChanged(self, value):
        self.progressBar.setValue(value)


class Worker(QtCore.QObject):
    
    @QtCore.pyqtSlot(str)
    def taskML(self, fileName, columnInput, sent_to_words, remove_stopwords, array_length, df, positive, negative, progressBarLabel, progressBar, hideProgNumLabel, runButton, alreadyPressed):        

        # Initialize start time to run script
        start_time = datetime.now()
        progressBarLabel.setText("Training Model...")
        
        data = df[columnInput.text().lower()].values.tolist()

        # Remove the new line character and single quotes
        data = [re.sub(r'\s+', ' ', str(sent)) for sent in data]
        data = [re.sub("\'", "", str(sent)) for sent in data]

        # Convert our data to a list of words. Now, data_words is a 2D array,
        # each index contains a list of words
        data_words = list(sent_to_words(data))

        # Remove the stop words
        data_words_nostops = remove_stopwords(data_words)
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

        progressBarLabel.setText("Finding risks...")
        for i in model.wv.most_similar(positive=positive, negative=negative, topn=similar_words_size):
            if len(i[0]) > 2:
                risks.append(i)

        highest = risks[0][1]

        progressBarLabel.setGeometry(QtCore.QRect(450, 835-103, 300, 25))
        progressBarLabel.setText("Tagging documents with risks...")
        risksWithDuplicates = defaultdict(list)
        riskCount = 0
        riskInDocCount = 0
        for document_number, i in enumerate(data_words_nostops):
            riskInDocConfirm = False
            for j in i:
                if len(j) > 2:
                    for k in risks:
                        if k[1] > highest/2 and k[0] in j and k[0][:3] == j[:3]:              
                            risksWithDuplicates[document_number].append(j)
                            riskCount+=1
                            riskInDocConfirm = True
                            break
            if riskInDocConfirm:
                riskInDocCount += 1

        risksWithDuplicates = dict(risksWithDuplicates)

        progressBarLabel.setGeometry(QtCore.QRect(500, 835-103, 300, 25))
        progressBarLabel.setText("Removing duplicates...")
        for key,value in risksWithDuplicates.items():
            # Removing duplicates
            newValue = list(dict.fromkeys(risksWithDuplicates[key]))
            newValue = {key: newValue}
            risksWithDuplicates.update(newValue)

        progressBarLabel.setText("Creating new file...")
        base = os.path.basename(fileName)
        name = os.path.splitext(base)[0]
        extension = os.path.splitext(base)[1]               
        df['Potential Risks'] = df.index.map(risksWithDuplicates)
        df['Machine Learning Statistics'] = ''
        df['Machine Learning Statistics'][0] = 'Total risks identefied in all documents: '+str(riskCount)
        df['Machine Learning Statistics'][1] = 'Average risk per document: '+str("{0:.2f}".format((riskCount/len(df.index))))
        df['Machine Learning Statistics'][2] = 'Percentage of documents that had no labeled risks: '+str("{0:.2f}".format((len(df.index)-riskInDocCount)/len(df.index))*100)+"%"


        cols = list(df.columns.values) 
        cols.pop(cols.index(columnInput.text()))
        cols.pop(cols.index('Potential Risks'))
        cols.pop(cols.index('Machine Learning Statistics'))
        df = df[cols+[columnInput.text(), 'Potential Risks', 'Machine Learning Statistics']]

        df.to_excel(name+"RisksIdentified"+extension.lower())
        move(os.path.dirname(os.path.realpath(__file__))+"\\\\"+name+"RisksIdentified"+extension, os.path.dirname(os.path.realpath(__file__))+"\\\\"+"files generated"+"\\\\"+name+"RisksIdentified"+extension)             
        
        progressBarLabel.setGeometry(QtCore.QRect(225, 815-103, 700, 70))
        hideProgNumLabel.hide()
        font = QtGui.QFont()
        font.setPointSize(10)
        progressBarLabel.setText("Risk words have been successfully tagged. A new excel file named \""+name+"RisksIdentified\""+" is generated to \""+os.path.dirname(os.path.realpath(__file__))+"\\files generated\\"+name+"RisksIdentified"+extension.lower()+"\" with a new column 'Potential Risks'")
        progressBarLabel.setFont(font)
        progressBar.hide()

        # Print time to run script
        print("Time taken to run script: ", datetime.now() - start_time)
        alreadyPressed[0] = False;

    @QtCore.pyqtSlot(str)
    def taskSearch(self, fileName, columnInput, sent_to_words, array_length, df, progressBarLabel, progressBar, hideProgNumLabel, runButton, searchInput, alreadyPressed):        

        # Initialize start time to run script
        start_time = datetime.now()
        progressBarLabel.setText("Searching for words...")

        data = df[columnInput.text().lower()].values.tolist()

        # Remove the new line character and single quotes
        data = [re.sub(r'\s+', ' ', str(sent)) for sent in data]
        data = [re.sub("\'", "", str(sent)) for sent in data]

        # Convert our data to a list of words. Now, data_words is a 2D array,
        # each index contains a list of words
        data_words = list(sent_to_words(data))

        progressBarLabel.setGeometry(QtCore.QRect(450, 835-103, 500, 25))
        progressBarLabel.setText("Tagging documents with found word...")
        wordsWithDuplicates = defaultdict(list)
        searchForWords = searchInput.toPlainText().replace('\n', '')
        searchForWords = searchInput.toPlainText().replace(" ", "").split(",")

        for document_number, document in enumerate(data_words):
            for word in document:
                for search in searchForWords:
                    if search == word:                        
                        wordsWithDuplicates[document_number].append(word)
                        

        progressBarLabel.setGeometry(QtCore.QRect(500, 835-103, 300, 25))
        progressBarLabel.setText("Removing duplicates...")
        wordsWithDuplicates = dict(wordsWithDuplicates)
        for key,value in wordsWithDuplicates.items():
            # Removing duplicates
            newValue = list(dict.fromkeys(wordsWithDuplicates[key]))
            newValue = {key: newValue}
            wordsWithDuplicates.update(newValue)        
        
        progressBarLabel.setText("Creating new file...")
        base = os.path.basename(fileName)
        name = os.path.splitext(base)[0]
        extension = os.path.splitext(base)[1]               
        df['Searched for Words'] = df.index.map(wordsWithDuplicates) # Removed duplicates

        cols = list(df.columns.values) 
        cols.pop(cols.index(columnInput.text()))
        cols.pop(cols.index('Searched for Words'))
        df = df[cols+[columnInput.text(), 'Searched for Words']]

        df.to_excel(name+"SearchedWords"+extension.lower())
        move(os.path.dirname(os.path.realpath(__file__))+"\\\\"+name+"SearchedWords"+extension, os.path.dirname(os.path.realpath(__file__))+"\\\\"+"files generated"+"\\\\"+name+"SearchedWords"+extension)

        progressBarLabel.setGeometry(QtCore.QRect(225, 815-103, 700, 70))
        hideProgNumLabel.hide()
        font = QtGui.QFont()
        font.setPointSize(10)
        progressBarLabel.setText("Searched words have been successfully tagged. A new excel file named \""+name+"SearchedWords\""+" is generated to \""+os.path.dirname(os.path.realpath(__file__))+"\\files generated\\"+name+"SearchedWords"+extension.lower()+"\" with a new column 'Searched for Words'")
        progressBarLabel.setFont(font)
        progressBar.hide()

        # Print time to run script
        print("Time taken to run script: ", datetime.now() - start_time)
        alreadyPressed[0] = False;

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