import pandas as pd
import requests
from bs4 import BeautifulSoup
import nltk
nltk.download('cmudict')
from nltk.tokenize import word_tokenize, sent_tokenize
from textblob import TextBlob

# Step 1: Data Extraction

# Load the URLs from the input.xlsx file into a pandas DataFrame
df_urls = pd.read_excel('Input.xlsx', sheet_name='Sheet1')

# Use a web scraping library to extract the HTML content of each URL
for index, row in df_urls.iterrows():
    url = row['URL']
    url_id = row['URL_ID']
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Parse the HTML content to extract only the article title and text
    article_title = soup.find('title').get_text()
    article_text = ''
    for paragraph in soup.find_all('p'):
        article_text += paragraph.get_text() + ' '

    # Save the extracted text in a text file with URL_ID as the filename
    with open(f'{url_id}.txt', 'w', encoding='utf-8') as f:
        f.write(f'{article_title}\n{article_text}')
        
import nltk
nltk.download('punkt')
from textblob import TextBlob
!pip install syllables
import syllables

# Step 2: Text Analysis

# Load the extracted text from each text file into a list
texts = []
for index, row in df_urls.iterrows():
    url_id = row['URL_ID']
    with open(f'{url_id}.txt', 'r', encoding='utf-8') as f:
        text = f.read()
        texts.append(text)

# Initialize a list to store the computed values for each variable
output = []

# Loop over the list of texts and perform the analysis for each text
for text in texts:
    # Tokenize the text into sentences and words using NLTK
    sentences = nltk.sent_tokenize(text)
    words = nltk.word_tokenize(text)

    # Calculate the number of words, syllables, and sentences in the text
    num_words = len(words)
    num_syllables = sum(syllables.estimate(word) for word in words)
    num_sentences = len(sentences)

    # Calculate the average sentence length and the average number of words per sentence
    avg_sentence_length = num_words / num_sentences
    avg_words_per_sentence = num_words / len(sentences)

    # Calculate the percentage of complex words
    num_complex_words = 0
    for word in words:
        if syllables.estimate(word) > 2:
            num_complex_words += 1
    percent_complex_words = (num_complex_words / num_words) * 100

    # Calculate the Fog index
    fog_index = 0.4 * (avg_words_per_sentence + percent_complex_words)

    # Count the number of personal pronouns in the text
    personal_pronouns = ['i', 'me', 'my', 'mine', 'you', 'your', 'yours', 'he', 'him', 'his', 'she', 'her', 'hers', 'it', 'its', 'we', 'us', 'our', 'ours', 'they', 'them', 'their', 'theirs']
    num_personal_pronouns = sum(word.lower() in personal_pronouns for word in words)

    # Calculate the positive score, negative score, polarity score, and subjectivity score using TextBlob
    blob = TextBlob(text)
    pos_score = sum(sentiment.polarity > 0 for sentiment in blob.sentences)
    neg_score = sum(sentiment.polarity < 0 for sentiment in blob.sentences)
    polarity_score = blob.sentiment.polarity
    subjectivity_score = blob.sentiment.subjectivity

    # Calculate the average word length
    sum_word_length = sum(len(word) for word in words)
    avg_word_length = sum_word_length / num_words

    # Store the computed values for each variable in a dictionary
    output_dict = {
        'num_words': num_words,
        'num_syllables': num_syllables,
        'num_sentences': num_sentences,
        'avg_sentence_length': avg_sentence_length,
        'avg_words_per_sentence': avg_words_per_sentence,
        'percent_complex_words': percent_complex_words,
        'fog_index': fog_index,
        'num_personal_pronouns': num_personal_pronouns,
        'pos_score': pos_score,
        'neg_score': neg_score,
        'polarity_score': polarity_score,
        'subjectivity_score': subjectivity_score,
        'avg_word_length': avg_word_length
    }

    # Append the output dictionary to the output list
    output.append(output_dict)
    
    # Create a new DataFrame from the output list of dictionaries
df_output = pd.DataFrame(output)

# Write the DataFrame to an Excel file
df_output.to_excel('Output.xlsx', index=False)

