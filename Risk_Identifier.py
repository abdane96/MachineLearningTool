import re
import numpy as np
import pandas as pd
import os
import nltk
from pprint import pprint
import warnings
from datetime import datetime
from pathlib import Path
'''import os
from gensim.models.wrappers import LdaMallet

os.environ['MALLET_HOME'] = 'C:/Users/qaraiea/Desktop/mallet-2.0.8'''''

# Gensim
import gensim
import gensim.corpora as corpora
from gensim.utils import simple_preprocess
from gensim.models import CoherenceModel
# Stop words
from nltk.corpus import stopwords

# spacy for lemmatization
# import spacy
'''
# Plotting tools
import pyLDAvis
import pyLDAvis.gensim  # don't skip this
import matplotlib.pyplot as plt
'''

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

'''
def make_bigrams(texts):
    return [bigram_mod[doc] for doc in texts]


def make_trigrams(texts):
    return [trigram_mod[bigram_mod[doc]] for doc in texts]


def lemmatization(texts, allowed_postags=['NOUN', 'ADJ', 'VERB', 'ADV']):
    texts_out = []
    for sent in texts:
        doc = nlp(" ".join(sent))
        texts_out.append([token.lemma_ for token in doc if token.pos_ in allowed_postags])
    return texts_out
'''

# Remove warning
warnings.filterwarnings("ignore", category=DeprecationWarning)

# Initialize stop words
stop_words = stopwords.words('english')
extended_stop_words = ['from', 're', 'edu', 'use', 'any', 'known']
stop_words.extend(extended_stop_words)

df = pd.read_csv("TestData.csv")

# Convert the data-frame(df) to a list
data = df.Comments.values.tolist()

# Remove the new line character and single quotes
data = [re.sub(r'\s+', ' ', str(sent)) for sent in data]
data = [re.sub("\'", "", str(sent)) for sent in data]

# Convert our data to a list of words, now data_words is a 2D array, each index contains a list of words
data_words = list(sent_to_words(data))
# Remove the stop words
data_words_nostops = remove_stopwords(data_words)


# Build the bigram and trigram models, the higher threshold, the fewer phrases.
# bigram = gensim.models.Phrases(data_words, min_count=5, threshold=100)
# trigram = gensim.models.Phrases(bigram[data_words], threshold=100)
# bigram_mod = gensim.models.phrases.Phraser(bigram)
# trigram_mod = gensim.models.phrases.Phraser(trigram)
#
# # Make a biagram
# data_words_bigrams = make_bigrams(data_words_nostops)
#
# # Initialize spacy 'en' model, keeping only tagger component (for efficiency)
# # python3 -m spacy download en
# nlp = spacy.load('en_core_web_sm', disable=['parser', 'ner'])
#
# # Do lemmatization keeping only noun, adj, vb, adv
# data_lemmatized = lemmatization(data_words_bigrams, allowed_postags=['NOUN', 'ADJ', 'VERB', 'ADV'])
#
# # Create Dictionary
# id2word = corpora.Dictionary(data_lemmatized)
#
# # Create Corpus
# texts = data_lemmatized
#
# # Term Document Frequency, the corpus format will follow: (word_id, word_frequency)
# corpus = [id2word.doc2bow(text) for text in texts]
#
# # Build LDA model
# lda_model = gensim.models.ldamodel.LdaModel(corpus=corpus, id2word=id2word, num_topics=7, random_state=100,
#                                             update_every=1, chunksize=50, passes=10, alpha='auto',
#                                             per_word_topics=True)
#
# pprint(lda_model.print_topics())

'''ldamallet = gensim.models.wrappers.LdaMallet("C:/Users/qaraiea/Desktop/mallet-2.0.8/bin/mallet", corpus=corpus,
                                             num_topics=10, id2word=id2word)'''


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

words_in_dictionary = set(nltk.corpus.words.words())

risks = []
positive = ['injuries', 'fail', 'dangerous', 'oil', 'incident']
negative = ['train', 'car', 'automobile', 'appliance', 'south', 'crossing', 'local', 'track', 'signal',
            'gate', 'bell', 'rail', 'westward']

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
    # print("Document number ", document_number, ":")
    for j in i:
        # print("Comparing word \"", j, "\" to:")
        if len(j) > 2:
            for k in risks:
                # print(k[0])
                if k[0] not in 'train' and k[1] > highest/2 and len(k[0]) > 2 and k[0] in j:
                    text_file.write("Document %i has the words %s\n" % (document_number, j))
                    break
text_file.close()
#   print(model.wv.most_similar(positive=positive, topn=0)[k][1])

# print(model.wv.most_similar(positive=positive, topn=0)[0][1])
# Print time to run script
print("Time taken to run script: ", datetime.now() - startTime)
