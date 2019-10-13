# -*- coding: utf-8 -*-
"""
Summary Generator 1.2
Created on Tue Aug  7 13:54:58 2018
UPDATE RELEASE NOTES: 
    Better memory management by efficient deletion
    Removes repeated duplicate processing
    Solves numeric characters at the end of word problem
    Resolves remaining key errors in tf_isf_scores
    Better function names
@author: X_Reaper_X
"""

import nltk
import xlrd
import math
import xlsxwriter
from nltk.tag import pos_tag
from nltk.corpus import PlaintextCorpusReader
from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer

#Define weight assigning macros for sentence feature sets

DEFINE_POSITION_LENGTH_WEIGHT       = 0.15
DEFINE_GRAMMATICAL_WEIGHT           = 0.39
DEFINE_THEMATIC_WEIGHT              = 0.28
DEFINE_CENTRALITY_SIMILARITY_WEIGHT = 0.18

#Define weight assigning macros for sentences
DEFINE_SENTENCE_LENGTH_WEIGHT       = 0.586
DEFINE_SENTENCE_POSITION_WEIGHT     = 0.414
DEFINE_SENTENCE_SIMILARITY_WEIGHT   = 0.465
DEFINE_SENTENCE_CENTRALITY_WEIGHT   = 0.533
DEFINE_SENTENCE_CUE_PHRASE_WEIGHT   = 0.421

#Define sentence-relevance macros
DEFINE_NUMBER_START_ANALYSIS        = 4
DEFINE_NUMBER_FINAL_SENTENCES_IMP   = 5
DEFINE_NUMBER_INITIAL_SENTENCES_IMP = 5

#Define weight assigning macros for words
DEFINE_WORD_ALPHANUMERIC_WEIGHT     = 0.289
DEFINE_WORD_NAMED_ENTITY_WEIGHT     = 0.344
DEFINE_WORD_IN_LEGAL_DICT_WEIGHT    = 0.588
DEFINE_ACTION_VERB_WEIGHT           = 0.267
DEFINE_TF_ISF_WEIGHT                = 0.143

#Define punctuation marks
DEFINE_PUNCTUATION_MARKS            = ['.',',',';','!','%','\n','\r','�',':','-','=','(',')','[',']','@','+','?',"'", '"']

corpus_root                         = 'C:\\Users\\X_Reaper_X\\Desktop\\Ready to test Oct20'    
dict_location                       = 'C:\\Users\\X_Reaper_X\\Downloads\\legal vocab.xlsx'
cue_phrases_location                = 'C:\\Users\\X_Reaper_X\\Downloads\\cue phrases.xlsx' 

#Regexp to read all .txt files located in the corpus root directory.  
file_ids = '.*.txt'    
                                              
corpus = PlaintextCorpusReader(corpus_root, file_ids)
porter = PorterStemmer()

#WORD PROCESSING GLOBAL VARIABLES

#Populate using tokenize()
#Reset using reset_tokens()
filtered_no_punct_word_tokens = []
stemmed_filtered_no_punct_word_tokens = []

#populate using make_ts_isf_database(word_tokens, sent_tokens)
#reset using reset_tf_isf_database()
tf_scores = {}                        
isf_scores = {}                       
tf_isf_scores = {}  

#populate using make_legal_dict() 
#same for each file in the corpus
legal_dict = {}
#poulate using make_cue_phrases_dict()
cue_phrases = {}

#populate using make_named_entity_database(word_tokens)
#reset using reset_named_entity_database()
named_entity = {}

#populate using make_title_words_database(sentence)
#reset using reset_title_words_database()
title_words = {}

#SENTENCE PROCESSING GLOBAL VARIABLES

#Populate using tokenize()
#Reser using reset_tokens()
sent_tokens = []
no_punct_sent_tokens = []
filtered_no_punct_sent_tokens = []
stemmed_filtered_no_punct_sent_tokens = []

#populate using make_sentence_scores_by_position_dict(sent_tokens)
#reset using reset_sentence_scores_by_position_dict()
sentence_scores_by_position = {}
sentence_scores_by_centrality = []
max_score = 0
"""
Returns a list of word tokens from a list of sentence tokens
INPUT- <lsit>sent_tokens
OUTPUT - <list>word_tokens
"""
def get_word_tokens_from_sent_tokens(sent_tokens):
    word_tokens = []
    for sent in sent_tokens:
        for word in sent.split():
            word_tokens.append(word)
    return word_tokens

#Resets global word and sentence variables
def reset_tokens():
    del filtered_no_punct_word_tokens[:]
    del stemmed_filtered_no_punct_word_tokens[:]
    del sent_tokens[:]
    del no_punct_sent_tokens[:]
    del filtered_no_punct_sent_tokens[:]
    del stemmed_filtered_no_punct_sent_tokens[:]
    return None

"""
Populate global variables
INPUT - <str>fileID
OUTPUT - <NoneType>
"""
def tokenize(fileID):
    global sent_tokens, stemmed_no_punct_sent_tokens, stemmed_filtered_no_punct_sent_tokens
    global stemmed_filtered_no_punct_word_tokens
    global max_score 
    
    reset_tokens()
    text = corpus.raw(fileID)
    
    #Populate global sentence variables
    sent_tokens = nltk.sent_tokenize(text)
    no_punct_sent_tokens = remove_sentences_punct(sent_tokens)
    filtered_no_punct_sent_tokens = remove_sentences_stopwords(no_punct_sent_tokens)
    stemmed_filtered_no_punct_sent_tokens = stem_sentences(filtered_no_punct_sent_tokens)
    
    #Populate global word variables
    filtered_no_punct_word_tokens = get_word_tokens_from_sent_tokens(filtered_no_punct_sent_tokens)
    stemmed_filtered_no_punct_word_tokens = get_word_tokens_from_sent_tokens(stemmed_filtered_no_punct_sent_tokens)
    
    #Populate global dictionaries
    make_legal_dict()
    make_cue_phrases_dict()
    make_named_entity_dict(filtered_no_punct_word_tokens)
    make_tf_isf_dict(stemmed_filtered_no_punct_word_tokens, stemmed_filtered_no_punct_sent_tokens)
    make_sentence_scores_by_position_dict(stemmed_filtered_no_punct_sent_tokens)
    make_title_words_dict(stemmed_filtered_no_punct_sent_tokens[0])
    
    max_score = max_centrality_score()
    
    
#SINGLE WORD PROCESSING FUNCTIONS

def remove_word_punct(word):
    no_punct_form = ''
    for letter in word:
        if letter not in DEFINE_PUNCTUATION_MARKS:
            no_punct_form += letter
    return no_punct_form

#CAUTION: porter stemmer returns word in lower case form
def stem_word(word, stemmer):
    return stemmer.stem(word)

#MULTIPLE WORDS PROCESSING FUNCTIONS
    

"""
Returns tf_isf_score of the provided word
INPUT - <str>stemmed_filtered_no_punct_word
OUTPUT - <double> tf_isf_score of the word
"""
def score_word_by_tf_isf(word):
    if word in tf_isf_scores.keys(): 
        return DEFINE_TF_ISF_WEIGHT * tf_isf_scores[word]
    return 0

"""
INPUT - <list>word_tokens
OUTPUT - <list>no_punct_tokens
"""
def remove_words_punct(tokens):
    no_punct_tokens = [] 
    for word in tokens:
        no_punct_form = remove_word_punct(word)
        if len(no_punct_form) != 0:
            no_punct_tokens.append(no_punct_form)
    return no_punct_tokens
"""
INPUT - <list>no_punct_word_tokens
OUTPUT - <list>filtered_no_punct_word_tokens
"""
def remove_stopwords(tokens):
    filtered_tokens = [word for word in tokens if word.lower() not in stopwords.words('english')]
    return filtered_tokens
"""
INPUT - <list>filtered_no_punct_tokens, <Stemmer>stemmer
OUTPUT - <list>lower_case_stemmed_filtered_no_punct_tokens
CAUTION - Stemmer returns lower case form
"""
def stem_tokens(tokens, stemmer):
    stemmed_tokens = [stem_word(word, stemmer) for word in tokens]
    return stemmed_tokens

#SINGLE SENTENCE PROCESING FUNCTIONS

def remove_sentence_punct(sent):
    return ' '.join(remove_words_punct(sent.split()))

def remove_sentence_stopwords(sent):
    filtered_sent_words = []
    for word in sent.split():
        if word.lower() not in stopwords.words('english'):
            filtered_sent_words.append(word)
    return ' '.join(filtered_sent_words)
"""
Returns stemmed sentence
INPUT - <str>filtered_no_punct_sentence
OUTPUT - <str>stemmed_filtered_no_punct_sentence
"""
def stem_sentence(sent):
    return ' '.join(stem_tokens(sent.split(), porter))

#MULTIPLE SENTENCES PROCESSING FUNCTIONS
"""
Returns a list of no punctuation forms of sentences
INPUT - <list>vanilla_sent_tokens
OUTPUT - <list>no_punct_sent_tokens
"""
def remove_sentences_punct(sent_tokens):
    no_punct_sent_tokens = [remove_sentence_punct(sent) for sent in sent_tokens]
    return no_punct_sent_tokens

"""
Returns a list of filtered sentences
INPUT - <list>no_punct_sent_tokens
OUTPUT - <list>filtered_sent_tokens
"""
def remove_sentences_stopwords(sent_tokens):
    filtered_sent_tokens = [remove_sentence_stopwords(sent) for sent in sent_tokens]
    return filtered_sent_tokens
"""
Returns a list of stemmed sentences
INPUT - <list>filtered_no_punct_sentences
OUTPUT - <lisy>stemmed_filtered_no_punct_sentences
"""
def stem_sentences(sent_tokens):
    stemmed_sent_tokens = [stem_sentence(sent) for sent in sent_tokens]
    return stemmed_sent_tokens

#GLOBAL DICTIONARY POPULATING FUNCTIONS

"""
Populates global variables  - tf_scores, isf_scores and tf_isf_scores - with lower case stemmed no punctuation tokens
INPUT - <list>stemmed_filtered_no_punct_word_tokens, <list>stemmed_filtered_no_punct_sent_tokens
OUTPUT - <NoneType>
"""
def make_tf_isf_dict(word_tokens, sent_tokens):
    reset_tf_isf_dict()
    word_occurrence_in_sent = 0
    for word in word_tokens:
        tf_scores[word] = word_tokens.count(word)/len(word_tokens)
        for sentence in sent_tokens:
            if sentence.find(word) != -1:
                word_occurrence_in_sent += 1
        isf_scores[word] = math.log(len(word_tokens)/(1+word_occurrence_in_sent))
        tf_isf_scores[word] = tf_scores[word]*isf_scores[word]
        word_occurrence_in_sent = 0
    normalize_tf_isf_scores()
    return None

def normalize_tf_isf_scores():
    max_tf_isf_score = max(tf_isf_scores.values())
    for word in tf_isf_scores.keys():
        tf_isf_scores[word] = tf_isf_scores[word]/max_tf_isf_score
    return None

#Resets global variables - tf_scores, isf_scores, tf_isf_scores.
def reset_tf_isf_dict():    
    tf_scores.clear()
    isf_scores.clear()
    tf_isf_scores.clear()
    return None

"""
Populates legal_dict with lower case stemmed form of the words provided in the excel file
INPUT - no arguments
OUTPUT - <NoneType>
"""
def make_legal_dict():
    workbook = xlrd.open_workbook(dict_location)
    sheet = workbook.sheet_by_index(0)
    for row in range(sheet.nrows):
        legal_dict[porter.stem(sheet.cell_value(row, 0))] = 1
    return None

def make_cue_phrases_dict():
    workbook = xlrd.open_workbook(cue_phrases_location)
    sheet = workbook.sheet_by_index(0)
    for row in range(sheet.nrows):
        cue_phrases[stem_sentence(remove_sentence_punct(sheet.cell_value(row, 0))).lower()] = 1
    return None

"""
Populates named_entity with lower case stemmed form of the words in word_tokens if pos_tagger tags a word as a proper or common noun and the length of the word is greater than 2
INPUT - <list>filtered_no_punct_word_tokens
OUTPUT - <NoneType>
"""
def make_named_entity_dict(word_tokens):
    reset_named_entity_dict()
    tagged_words = pos_tag(word_tokens)
    for word, pos in tagged_words:
        if len(word) > 2:
            if pos == 'NNP':
                named_entity[porter.stem(word)] = 1
    return None

"""
Populates global variable - <dict>sentence_scores_by_positio - with relevant initial and final sentences and their respective scores
INPUT - <list>stemmed_filtered_no_punct_sent_tokens
OUTPUT - <NoneType>
"""
def make_sentence_scores_by_position_dict(sent_tokens):
    counter = 0
    for n in range(len(sent_tokens)):
        if counter < DEFINE_NUMBER_INITIAL_SENTENCES_IMP:
            if len(sent_tokens[n].split()) > 5:
                counter += 1
                sentence_scores_by_position[sent_tokens[n]] = 1/counter
        else:
            break
    counter = 0
    for n in range(len(sent_tokens)-1, 0, -1):
        if counter < DEFINE_NUMBER_FINAL_SENTENCES_IMP:
            if len(sent_tokens[n].split()) > 5:
                counter += 1
                sentence_scores_by_position[sent_tokens[n]] = 1/counter
        else:
            break
    return None

"""
Populates global variable - <dict>title_words - with lower case no punctuation stemmed form of the words in the first sentence in case of legal documents
INPUT - <str>stemmed_filtered_no_punct_sent: first word of the document
OUTPUT - <NoneType>
"""
def make_title_words_dict(sent):
    reset_title_words_dict()
    for word in sent.split():
        title_words[word] = 1
    return None

"""
Returns a non-zero score if word is in legal vocab otherwise 0
INPUT - <str>lower case stemmed form of the word
OUTPUT - <double>score
"""
def score_word_if_legal(word):
    if word in legal_dict.keys():
        return DEFINE_WORD_IN_LEGAL_DICT_WEIGHT * legal_dict[word]
    return 0.0

def score_sentence_if_has_cue_phrase(sent):
    for phrase in cue_phrases.keys():
        if sent.find(phrase) != -1:
            return DEFINE_SENTENCE_CUE_PHRASE_WEIGHT * cue_phrases[phrase]
    return 0

"""
Returns a non-zero score if word has a numeric character otherwise 0
INPUT - <str>word
OUTPUT - <double>score
"""
def score_word_if_has_numeric(word):
    word_score_by_character_type = 0.0
    for letter in word:
        if letter.isnumeric():
            word_score_by_character_type = 1
            break
    return DEFINE_WORD_ALPHANUMERIC_WEIGHT * word_score_by_character_type



#Resets global variable - named_entity
def reset_named_entity_dict():
    named_entity.clear()
    return None

"""
Returns a non-zero score if provided word is a named entity
INPUT - <str>lower case stemmed form of the word
OUTPUT - <double>score
"""
def score_word_if_named_entity(word):
    if word in named_entity.keys():
        return DEFINE_WORD_NAMED_ENTITY_WEIGHT * named_entity[word]
    return 0.0

#SENTENCE PROCESSING


#Resets global variable - <dict>sentence_scores_by_position 
def reset_sentence_scores_by_position_dict():
    sentence_scores_by_position.clear()
    return None

def score_sentence_by_position(sent):
    if sent in sentence_scores_by_position.keys():
       return DEFINE_SENTENCE_POSITION_WEIGHT * sentence_scores_by_position[sent]
    return 0.0

def max_centrality_score():
    global sentence_scores_by_centrality
    for sent in stemmed_filtered_no_punct_word_tokens:
        sentence_scores_by_centrality.append(score_sentence_by_centrality(sent))
    return max(sentence_scores_by_centrality)

#INPUT stemmed_filtered_no_punct_sentence
def score_sentence_by_centrality(sent):
    sent_score = 0
    for word in sent.split():
        if word not in stopwords.words('english'):
            sent_score += stemmed_filtered_no_punct_word_tokens.count(word)
    return DEFINE_SENTENCE_CENTRALITY_WEIGHT * sent_score

def normalized_score_sentence_by_centrality(sent):
    global max_score
    return score_sentence_by_centrality(sent)/max_score
"""
Returns a score based on the length of the sentence
INPUT - <str>sentence
OUTPUT - <double>score
"""
def score_sentence_by_len(sentence):
    sentence_score_by_len = 0
    if len(sentence.split()) < 3:
        sentence_score_by_len += 0.13
    elif len(sentence.split()) < 6:
        sentence_score_by_len += 0.25
    elif len(sentence.split()) < 12:
        sentence_score_by_len += 0.50
    elif len(sentence.split()) < 30:
        sentence_score_by_len += 1
    else:
        sentence_score_by_len += 0.13
    return DEFINE_SENTENCE_LENGTH_WEIGHT * sentence_score_by_len



#Resets global variable - <dict>title_words
def reset_title_words_dict():
    title_words.clear()
    return None   
"""
Returns a non zero score if the provided word is in the title of the text
INPUT - <str>stemmed_filtered_no_punct_word
OUTPUT - <double>score
"""
def score_word_if_in_title_words(word):
    if word in title_words.keys():
        return DEFINE_SENTENCE_SIMILARITY_WEIGHT * title_words[word]
    return 0

def score_word(word):
    tf_isf_score = score_word_by_tf_isf(word)
    legal_score = score_word_if_legal(word)
    char_score = score_word_if_has_numeric(word)
    named_entity_score = score_word_if_named_entity(word)
    title_score = score_word_if_in_title_words(word)
    return tf_isf_score + legal_score + char_score + named_entity_score + title_score

def score_sentence_by_tf_isf(sent):
    tf_isf_score = 0
    for word in sent.split():
        tf_isf_score += score_word_by_tf_isf(word)
    return tf_isf_score

def score_sentence_by_legal_dict(sent):
    legal_score = 0
    for word in sent.split():
        legal_score += score_word_if_legal(word)
    return legal_score

def score_sentence_by_numbers_dates(sent):
    char_score = 0
    for word in sent.split():
        char_score += score_word_if_has_numeric(word)
    return char_score

def score_sentence_by_named_entity(sent):
    named_entity_score = 0
    for word in sent.split():
        named_entity_score += score_word_if_named_entity(word)
    return named_entity_score

def score_sentence_by_title_words(sent):
    similarity_score = 0
    for word in sent.split():
        similarity_score += score_word_if_in_title_words(word)
    return similarity_score

def score_sentence_by_words(sent):
    overall_score = 0
    for word in sent.split():
        overall_score += score_word(word)
    return overall_score

def score_sentence_by_action_verbs(sent):
    action_verb_score = 0
    tagged_words = pos_tag(sent.split())
    for word, pos in tagged_words:
        if len(word) > 2:
            if pos == 'VB':
                action_verb_score +=1 
    return DEFINE_ACTION_VERB_WEIGHT * action_verb_score
    
def score_sentence(sent):
    length_score = score_sentence_by_len(sent)
    position_score = score_sentence_by_position(sent)
    
    tf_isf_score = score_sentence_by_tf_isf(sent)
    named_entity_score = score_sentence_by_named_entity(sent)
    char_score = score_sentence_by_numbers_dates(sent)
    verb_score = score_sentence_by_action_verbs(sent)
    
    legal_score = score_sentence_by_legal_dict(sent)
    cue_phrase_score = score_sentence_if_has_cue_phrase(sent)
    
    centrality_score = normalized_score_sentence_by_centrality(sent)
    similarity_score = score_sentence_by_title_words(sent)
    
    
    final_score =  DEFINE_POSITION_LENGTH_WEIGHT*(length_score + position_score) + DEFINE_GRAMMATICAL_WEIGHT*(tf_isf_score + named_entity_score + char_score + verb_score) + DEFINE_THEMATIC_WEIGHT*(legal_score + cue_phrase_score) + DEFINE_CENTRALITY_SIMILARITY_WEIGHT*(centrality_score + similarity_score)
                   
    return final_score

def gen_summary():
    for fileID in corpus.fileids():
        tokenize(fileID)
        sent_count = 0
        for n in range(len(sent_tokens)):
            if score_sentence(stemmed_filtered_no_punct_sent_tokens[n]) > 3.5:
                print(sent_tokens[n])
                sent_count += 1
        print(sent_count)
        print('\n\n\n\n\n')
        print('************************************************')
        


def write_feature_scores_to_excel_sheet():
    workbook = xlsxwriter.Workbook('Feature Scores.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(0,0, 'Sent No')
    worksheet.write(0,1, 'Legal Score')
    worksheet.write(0,2, 'TF ISF Score')
    worksheet.write(0,3, 'Char Score')
    worksheet.write(0,4, 'Named Entity Score')
    worksheet.write(0,5, 'Title Score')
    worksheet.write(0,6, 'Length Score')
    worksheet.write(0,7, 'Position Score')
    worksheet.write(0,8, 'Cue Phrase Score')
    worksheet.write(0,9, 'Centrality Score')
    worksheet.write(0, 10, 'Verb Score')
    
    
    for fileID in corpus.fileids():
        tokenize(fileID)
    row = 1
    for sent in stemmed_filtered_no_punct_sent_tokens:
        legal_score = 0
        tf_isf_score = 0
        char_score = 0
        named_entity_score = 0
        title_score = 0
        length_score = score_sentence_by_len(sent)
        position_score = score_sentence_by_position(sent)
        cue_phrase_score = score_sentence_if_has_cue_phrase(sent)
        centrality_score = normalized_score_sentence_by_centrality(sent)
        verb_score = score_sentence_by_action_verbs(sent)
        for word in sent.split(' '):
            tf_isf_score += score_word_by_tf_isf(word)
            legal_score += score_word_if_legal(word)
            char_score += score_word_if_has_numeric(word)
            named_entity_score += score_word_if_named_entity(word)
            title_score += score_word_if_in_title_words(word)
        row += 1
        col = 0
        worksheet.write(row,col, row)
        worksheet.write(row,col+1, legal_score)
        worksheet.write(row,col+2, tf_isf_score)
        worksheet.write(row,col+3, char_score)
        worksheet.write(row,col+4, named_entity_score)
        worksheet.write(row,col+5, title_score)
        worksheet.write(row,col+6, length_score)
        worksheet.write(row,col+7, position_score)
        worksheet.write(row,col+8, cue_phrase_score)
        worksheet.write(row,col+9, centrality_score)
        worksheet.write(row,col+10, verb_score)
        
gen_summary()