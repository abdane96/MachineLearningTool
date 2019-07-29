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
    prompt = filedialog.askopenfilename()
    filename.set(os.path.basename(prompt))
    if prompt is not "":
        filename.set('Selected file: '+filename.get())


def column_in_columns(column_names, text):
    for i in column_names:
        if text == i:
            return True
    return False


# Remove warning
warnings.filterwarnings("ignore", category=DeprecationWarning)
# Initialize stop words
stop_words = stopwords.words('english')
extended_stop_words = ['from', 're', 'edu', 'use', 'any', 'known']
stop_words.extend(extended_stop_words)


def run():
    df = pd.read_excel(prompt)
    df.columns = map(str.lower, df.columns)

    # Convert the data-frame(df) to a list
    if columnNameField.get() == "":
        messagebox.showinfo("Error", "Please provide a column name!")
    elif not column_in_columns(df.columns, columnNameField.get().lower()):
        messagebox.showinfo("Error", "The column name you provided does not exist!")
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
        positive = ['injuries', 'fail', 'dangerous', 'oil', 'incident', 'struck',
                    'hit', 'derail', 'incorrect', 'collision', 'fatal', 'leaking', 'tamper', 'emergency',
                    'gas', 'collided', 'damage', 'risk', 'broken', 'break']
        negative = ['goods', 'calgary', 'train', 'car', 'automobile', 'appliance', 'south', 'mileage',
                    'crossing', 'local', 'track', 'signal', 'gate', 'bell', 'rail', 'westward']

        similar_words_size = array_length(model.wv.most_similar(positive=positive, negative=negative, topn=0))
        text_file = open("Output.txt", "w")

        for i in model.wv.most_similar(positive=positive, negative=negative, topn=similar_words_size):
            if len(i[0]) > 2:
                # risks.append(i[0])
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

        average = sum/count

        for i in data_words_nostops:
            document_number += 1
            for j in i:
                if len(j) > 2:
                    for k in risks:
                        if k[0] not in 'train' and k[1] > highest/2 and len(k[0]) > 2 and k[0] in j:
                            text_file.write("Document %i has the words %s\n" % (document_number, j))
                            break
        text_file.close()
        # print(model.wv.most_similar(positive=positive, topn=0)[k][1])

        # print(model.wv.most_similar(positive=positive, topn=0)[0][1])


root = Tk()
root.title("Machine learning risk identifier")

filename = StringVar()
prompt = ""
browseBtn = Button(text="Browse", command=browse_button)
browseLabel = Label(master=root, textvariable=filename)
title = Label(master=root, text="Identify risks in a document", font=("Helvetica", 22), pady=20)
title.grid(row=0, column=1)
browseLabel.grid(row=1, column=1)
browseBtn.grid(row=1, column=0, padx=20, pady=20)
columnNameLabel = Label(root, text="Column Name").grid(row=3, column=0, ipadx=10)
columnNameField = Entry(root, width=45)
columnNameField.grid(row=3, column=1)
positiveWordsLabel = Label(root, text="Positive words (separate by comma)").grid(row=4, column=0, ipady=20, ipadx=10)
positiveWordsField = Entry(root, width=45)
positiveWordsField.grid(row=4, column=1)
runBtn = Button(text="Run", command=run)
runBtn.grid(row=1, column=2, padx=20, pady=10)


mainloop()

# Print time to run script
print("Time taken to run script: ", datetime.now() - startTime)
