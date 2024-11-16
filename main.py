import pandas as pd
import requests
from bs4 import BeautifulSoup
import nltk
import re

# Download nltk data if not already downloaded
nltk.download('punkt')
nltk.download('stopwords')
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize

# Load stop words and positive/negative words
stop_words = set(stopwords.words("english"))
positive_words = set(open('MasterDictionary/positive-words.txt').read().splitlines())
negative_words = set(open('MasterDictionary/negative-words.txt').read().splitlines())

# Column order required for output
output_columns = [
    "URL_ID", "URL", "POSITIVE SCORE", "NEGATIVE SCORE", "POLARITY SCORE", 
    "SUBJECTIVITY SCORE", "AVG SENTENCE LENGTH", "PERCENTAGE OF COMPLEX WORDS", 
    "FOG INDEX", "AVG NUMBER OF WORDS PER SENTENCE", "COMPLEX WORD COUNT", 
    "WORD COUNT", "SYLLABLE PER WORD", "PERSONAL PRONOUNS", "AVG WORD LENGTH"
]

# Function to fetch and extract text from URL
def fetch_article_text(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    title = soup.find('h1').get_text(strip=True) if soup.find('h1') else ''  # Handling missing titles
    paragraphs = soup.find_all('p')
    text = ' '.join([para.get_text(strip=True) for para in paragraphs])
    
    return title + "\n\n" + text

# Text Analysis functions
def compute_sentiment_scores(text):
    words = word_tokenize(text.lower())
    words = [word for word in words if word.isalpha() and word not in stop_words]

    positive_score = sum(1 for word in words if word in positive_words)
    negative_score = sum(1 for word in words if word in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(words) + 0.000001)
    
    return positive_score, negative_score, polarity_score, subjectivity_score

def compute_readability_scores(text):
    sentences = sent_tokenize(text)
    words = word_tokenize(text.lower())
    words = [word for word in words if word.isalpha()]
    
    avg_sentence_length = len(words) / len(sentences) if sentences else 0
    complex_word_count = sum(1 for word in words if count_syllables(word) > 2)
    percentage_complex_words = complex_word_count / len(words) if words else 0
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words) if avg_sentence_length > 0 else 0
    
    return avg_sentence_length, percentage_complex_words, fog_index, complex_word_count, len(words)

def count_syllables(word):
    word = word.lower()
    syllable_count = len(re.findall(r'[aeiouy]+', word))
    if word.endswith(("es", "ed")):
        syllable_count -= 1
    return max(1, syllable_count)

def analyze_text(text):
    positive_score, negative_score, polarity_score, subjectivity_score = compute_sentiment_scores(text)
    avg_sentence_length, percentage_complex_words, fog_index, complex_word_count, word_count = compute_readability_scores(text)
    syllable_per_word = sum(count_syllables(word) for word in word_tokenize(text) if word.isalpha()) / word_count if word_count > 0 else 0
    personal_pronouns = len(re.findall(r'\b(I|we|my|ours|us)\b', text, re.I))
    avg_word_length = sum(len(word) for word in word_tokenize(text) if word.isalpha()) / word_count if word_count > 0 else 0

    return {
        "POSITIVE SCORE": positive_score,
        "NEGATIVE SCORE": negative_score,
        "POLARITY SCORE": polarity_score,
        "SUBJECTIVITY SCORE": subjectivity_score,
        "AVG SENTENCE LENGTH": avg_sentence_length,
        "PERCENTAGE OF COMPLEX WORDS": percentage_complex_words,
        "FOG INDEX": fog_index,
        "AVG NUMBER OF WORDS PER SENTENCE": avg_sentence_length,
        "COMPLEX WORD COUNT": complex_word_count,
        "WORD COUNT": word_count,
        "SYLLABLE PER WORD": syllable_per_word,
        "PERSONAL PRONOUNS": personal_pronouns,
        "AVG WORD LENGTH": avg_word_length,
    }

# Load input URLs from Excel and process
input_file = 'Input.xlsx'
output_file = 'Output Data Structure.xlsx'
data = pd.read_excel(input_file)

# Create an empty list to store results
results = []

for _, row in data.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    try:
        print(f"Processing URL_ID: {url_id}, URL: {url}")  # Debugging output
        article_text = fetch_article_text(url)
        
        # Perform text analysis
        analysis_results = analyze_text(article_text)
        result_row = {**row.to_dict(), **analysis_results}
        results.append(result_row)

    except Exception as e:
        print(f"Error processing URL {url_id}: {e}")

# Convert results to DataFrame and ensure it matches output structure
output_df = pd.DataFrame(results, columns=output_columns)

# Save the output to a new Excel file with the defined structure
output_df.to_excel(output_file, index=False)
print("Data successfully saved to Excel.")
