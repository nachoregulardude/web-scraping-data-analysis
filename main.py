import requests
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from bs4 import BeautifulSoup
import pandas as pd

def getLinks():
    # takes in the Excel file and returns the page links as a list
    df = pd.read_excel('Input.xlsx')
    data = pd.DataFrame(df, columns=['URL'])
    return data['URL'].tolist()

def getData(url, headers):
    # Gets the blog data from the url
    r = requests.get(url, headers=headers)
    print(f'Get request status code: {r.status_code}')
    sp = BeautifulSoup(r.content, features='lxml')
    data = sp.find("div", class_="td-post-content")
    # remove the data about authors
    for item in data("pre"):
        item.decompose()
    return data.text

def parseData(data, lmdict, url):
    # parses the data and returns the sentiment values
    # Self explanatory variable names
    positiveScore = 0
    negativeScore = 0
    wordLen = 0
    filtered_sentence =''
    #Create a set of stop words to ignore
    stop_words = set(stopwords.words('english'))
    data = data.replace('\n', ' ')
    # this section will remove the punctuation
    punctuation = '''!()-[]{};:'"\,<>./?@#$%^&*_~'''
    data_w_punc = data
    for ele in data:
        if ele in punctuation:
            data = data.replace(ele, "")

    # Create a list of words
    data_list = data.split(' ')

    # This section will remove empty spaces and unwanted seperators
    while '' in data_list:
        data_list.remove('')
    while '–' in data_list:
        data_list.remove('–')

    wordLen = len(data_list)
    for word in data_list:
        tokenized = word_tokenize(word)
        if word in stop_words:
            continue
        filtered_sentence += word + ' '
    for word in filtered_sentence.split(" "):
        if word.lower() in lmdict['Positive']:
            positiveScore += 1
        elif word.lower() in lmdict['Negative']:
            negativeScore += 1

    # Calculation of different values
    polarityScore = (positiveScore - negativeScore)/ ((positiveScore + negativeScore) + 0.000001) 
    subjectivityScore = (positiveScore + negativeScore)/ ((len(filtered_sentence)) + 0.000001)
    try:
        avgSentenceLeng = wordLen/(data_w_punc.count('.'))
    except ZeroDivisionError:
        avgSentenceLeng = 1

    complex_count = 0 
    syllable_count = 0

    for word in data_list:
        syllable_count += syllable_counter(word)
        if syllable_counter(word)>2:
            complex_count += 1

    perc_of_complex_words = complex_count/wordLen
    fog_index = 0.4 * (avgSentenceLeng + perc_of_complex_words)
    avg_words_per_sentences = wordLen/(data_w_punc.count('.'))
    avg_word_len = len(data)/len(data_list)
    # Creating a dictionary with the required values
    result = {
            'URL': url,
            'POSITIVE SCORE': positiveScore,
            'NEGATIVE SCORE': negativeScore,
            'POLARITY SCORE': polarityScore,
            'SUBJECTIVITY SCORE': subjectivityScore,
            'AVG SENTENCE LENGTH': avgSentenceLeng,
            'PERCENTAGE OF COMPLEX WORDS': perc_of_complex_words,
            'FOG INDEX': fog_index,
            'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentences,
            'COMPLEX WORD COUNT': complex_count,
            'WORD COUNT': len(filtered_sentence.split(' ')),
            'SYLLABLE PER WORD': syllable_count,
            'PERSONAL PRONOUNS': personal_pronoun_counter(data_list),
            'AVG WORD LENGTH': avg_word_len

            }
    return result

def personal_pronoun_counter(data_list):
    # Given a list of strings, returns the number of personal pronouns
    words = ['I', 'we', 'We', 'my', 'My', 'ours', 'Ours', 'us', 'Us']
    count = 0
    for pronoun in words:
        count += data_list.count(pronoun)
        
    return count

def syllable_counter(word):
    # Given a word, returns the number of vowels
    vowels = 'aeiou'
    syllable_count = 0
    for ele in word:
        if ele in vowels and 'es' not in word and 'ed' not in word:
            syllable_count += 1
    return syllable_count

def write_to_excel(result):
    df = pd.DataFrame(result)
    df.to_excel('Output Data Structure.xlsx', index=False)



def createMaster():
    # Creating a MasterDictionary with words and their sentiment values
    masterDF = pd.read_csv(r'Loughran-McDonald_MasterDictionary_1993-2021.csv')
    words = pd.DataFrame(masterDF, columns=['Word', 'Negative', 'Positive'])
    lmdict = {'Positive': list(), 'Negative': list()}
    for i in words.index:
        if (words['Negative'][i] > 0):
            lmdict['Negative'].append(words['Word'][i].lower())
        elif (words['Positive'][i] > 0):
            lmdict['Positive'].append(words['Word'][i].lower())
    return lmdict

def main():
    print("Collecting the links from the CSV file...")
    urls = getLinks()
    print("Creating a sentiment dictionary...")
    lmdict = createMaster()
    # User agent information
    headers = {'User-Agent' : 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.74 Safari/537.36'}
    results = []
    for url in urls:
        print(f"Working on {url}")
        data = getData(url, headers)
        result = parseData(data, lmdict, url)
        results.append(result)

    write_to_excel(results)
    print('Saved to the output excel file.')


if __name__ == "__main__":
    main()
