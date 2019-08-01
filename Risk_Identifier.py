import re
import numpy as np
import pandas as pd
import os
import nltk
from pprint import pprint
import warnings
from datetime import datetime
from pathlib import Path
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

# Gensim
import gensim
import gensim.corpora as corpora
from gensim.utils import simple_preprocess
from gensim.models import CoherenceModel

# Stop words
from nltk.corpus import stopwords

# Initialize start time to run script
startTime = datetime.now()


def sent_to_words(sentences):
    # Tokenize each sentence into a list of words and remove unwanted characters
    for sentence in sentences:
        yield(gensim.utils.simple_preprocess(str(sentence), deacc=True))


def array_length(array):
    count = 0
    for i in array:
        count += 1
    return count


def remove_stopwords(texts):
    # Removes stopwords in a text
    return [[word for word in simple_preprocess(str(doc)) if word not in stop_words] for doc in texts]


def browse_button():
    # Allow user to select a file and store the directory in global var
    # called folder_path
    global prompt
    prompt = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
    filename.set(os.path.basename(prompt))
    if prompt is not "":
        filename.set('Selected file: '+filename.get())


def column_in_columns(column_names, text):
    for i in column_names:
        if text == i:
            return True
    return False


def check_box(var, field):
    if var.get() == 1:
        field.config(state='disabled')
    else:
        field.config(state='normal')


# Remove warning
warnings.filterwarnings("ignore", category=DeprecationWarning)
# Initialize stop words
stop_words = stopwords.words('english')
extended_stop_words = ['from', 're', 'edu', 'use', 'any', 'known']
stop_words.extend(extended_stop_words)


def run():
    if prompt is "":
        messagebox.showinfo("Error", "Please select a document first!")
    else:
        xl = pd.ExcelFile(prompt)
        if SheetNameField.get() == "":
            return messagebox.showinfo("Error", "Please provide a sheet name!")
        elif not column_in_columns(map(str.lower, xl.sheet_names), SheetNameField.get().lower()):
            return messagebox.showinfo("Error", "The sheet name you provided does not exist!")
        else:
            df = pd.read_excel(prompt)
            df.columns = map(str.lower, df.columns)

            # Convert the data-frame(df) to a list
            if columnNameField.get() == "":
                return messagebox.showinfo("Error", "Please provide a column name!")
            elif not column_in_columns(df.columns, columnNameField.get().lower()):
                return messagebox.showinfo("Error", "The column name you provided does not exist!")
            else:
                data = df[columnNameField.get().lower()].values.tolist()

                # Remove the new line character and single quotes
                data = [re.sub(r'\s+', ' ', str(sent)) for sent in data]
                data = [re.sub("\'", "", str(sent)) for sent in data]

                # Convert our data to a list of words. Now, data_words is a 2D array, each index contains a list of words
                data_words = list(sent_to_words(data))
                # Remove the stop words
                data_words_nostops = remove_stopwords(data_words)

                model = gensim.models.Word2Vec(
                        data_words_nostops,
                        size=300,
                        alpha=0.03,
                        window=2,
                        min_count=2,
                        min_alpha=0.0007,
                        negative=20,
                        workers=10)
                model.train(data_words_nostops, total_examples=len(data_words_nostops), epochs=10)

                risks = []

                if pos_var.get() == 1 and neg_var.get() == 1:
                    positive = ['injuries', 'fail', 'dangerous', 'oil', 'incident', 'struck',
                                'hit', 'derail', 'incorrect', 'collision', 'fatal', 'leaking', 'tamper', 'emergency',
                                'gas', 'collided', 'damage', 'risk', 'broken', 'break']

                    negative = ['goods', 'calgary', 'train', 'car', 'automobile', 'appliance', 'south', 'mileage',
                                'crossing', 'local', 'track', 'signal', 'gate', 'bell', 'rail', 'westward']

                elif pos_var.get() == 0 and neg_var.get() == 1:
                    positive_words = positiveWordsField.get().replace(" ", "").split(",")
                    positive = positive_words
                    if positive[0] is '':
                        return messagebox.showinfo("Error", "Positive/Negative words can't be empty, write something"
                                                     " or use default values!")

                    negative = ['goods', 'calgary', 'train', 'car', 'automobile', 'appliance', 'south', 'mileage',
                                'crossing', 'local', 'track', 'signal', 'gate', 'bell', 'rail', 'westward']
                elif pos_var.get() == 1 and neg_var.get() == 0:
                    negative_words = positiveWordsField.get().replace(" ", "").split(",")
                    negative = negative_words
                    if negative[0] is '':
                        return messagebox.showinfo("Error", "Positive/Negative words can't be empty, write something"
                                                     " or use default values!")
                    positive = ['injuries', 'fail', 'dangerous', 'oil', 'incident', 'struck',
                                'hit', 'derail', 'incorrect', 'collision', 'fatal', 'leaking', 'tamper', 'emergency',
                                'gas', 'collided', 'damage', 'risk', 'broken', 'break']
                else:
                    positive_words = positiveWordsField.get().replace(" ", "").split(",")
                    positive = positive_words
                    negative_words = positiveWordsField.get().replace(" ", "").split(",")
                    negative = negative_words
                    if positive[0] is '' or negative[0] is '':
                        return messagebox.showinfo("Error", "Positive/Negative words can't be empty, write something"
                                                     " or use default values!")

                similar_words_size = array_length(model.wv.most_similar(positive=positive, negative=negative, topn=0))
                text_file = open("Output.txt", "w")

                for i in model.wv.most_similar(positive=positive, negative=negative, topn=similar_words_size):
                    if len(i[0]) > 2:
                        risks.append(i)

                # print(risks)

                document_number = 0
                count = 0
                sum = 0

                highest = risks[0][1]

                for j in risks:
                    if j[1] > 0:
                        sum += j[1]
                        count += 1

                for i in data_words_nostops:
                    document_number += 1
                    for j in i:
                        if len(j) > 2:
                            for k in risks:
                                if k[1] > highest/2 and len(k[0]) > 2 and k[0] in j:
                                    text_file.write("Document %i has the words %s\n" % (document_number, j))
                                    break
                text_file.close()
                messagebox.showinfo("Success!", "Risk words have been successfully tagged in a new column 'Potential Risks'")


# Root
root = Tk()
root.title("Machine learning risk identifier")

# File info
filename = StringVar()
prompt = ""

# Header title
title = Label(master=root, text="Identify risks in a document", font=("Helvetica", 22), pady=20)
title.grid(row=0, column=0)

# Browse Button
browseBtn = Button(text="Browse", command=browse_button)
browseLabel = Label(master=root, textvariable=filename)
browseLabel.grid(row=1, column=1)
browseBtn.grid(row=1, column=0, padx=20, pady=20)

# Run Button
runBtn = Button(text="Run", command=run)
runBtn.grid(row=1, column=2, padx=20, pady=10)

# Sheet Number
SheetNameLabel = Label(root, text="Sheet Name").grid(row=3, column=0, ipady=20, ipadx=10)
SheetNameField = Entry(root, width=45)
SheetNameField.grid(row=3, column=1)

# Column Name
columnNameLabel = Label(root, text="Column Name").grid(row=4, column=0, ipadx=10)
columnNameField = Entry(root, width=45)
columnNameField.grid(row=4, column=1)

# Positive Words
positiveWordsLabel = Label(root, text="Positive words separated by commas "
                                      "(words must exits in documents)").grid(row=5, column=0, ipady=30, ipadx=10)
positiveWordsField = Entry(root, width=45)
positiveWordsField.grid(row=5, column=1)
positiveWordsField.config(state='disabled')
pos_var = IntVar()
defaultPositiveWords = Checkbutton(root, text="Default Positive Words", variable=pos_var,
                                   command=lambda: check_box(pos_var, positiveWordsField))
defaultPositiveWords.grid(row=5, column=2)
defaultPositiveWords.toggle()

# Negative Words
negativeWordsLabel = Label(root, text="Negative words separated by commas "
                                      "(words must exits in documents)").grid(row=6, column=0, ipady=20, ipadx=10)
negativeWordsField = Entry(root, width=45)
negativeWordsField.grid(row=6, column=1)
negativeWordsField.config(state='disabled')
neg_var = IntVar()
defaultNegativeWords = Checkbutton(root, text="Default Negative Words", variable=neg_var,
                                   command=lambda: check_box(neg_var, negativeWordsField))
defaultNegativeWords.grid(row=6, column=2)
defaultNegativeWords.toggle()


mainloop()

# Print time to run script
print("Time taken to run script: ", datetime.now() - startTime)
