import pandas as pd
from nltk.tokenize import word_tokenize
from nltk.sentiment import SentimentIntensityAnalyzer
from tqdm import tqdm

# Download NLTK resources
import nltk
nltk.download('punkt')
nltk.download('vader_lexicon')

# Read the Excel file into a DataFrame
file_path = "E:/Study Abroad/Semester 01/MANG6540 - Integrated Marketing/Final Assessment/Sentiment Analysis/Sentiment Data.xlsx"
df = pd.read_excel(file_path)

# Tokenize the "Review" column
df['tokens'] = df['Review'].apply(word_tokenize)

# Sentiment Scoring using nltk SentimentIntensityAnalyzer for each token
sia = SentimentIntensityAnalyzer()

def get_sentiment_score(tokens):
    scores = [sia.polarity_scores(token)['compound'] for token in tokens]
    return scores

df['sentiment_scores'] = df['tokens'].apply(get_sentiment_score)

# Flatten the list of sentiment scores
df_flat = df.explode('sentiment_scores')

# Combine the flattened DataFrame with the original DataFrame
df_combined = pd.concat([df, df_flat['sentiment_scores'].apply(pd.Series)], axis=1)

# Save the result to an Excel file
output_path = "E:/Study Abroad/Semester 01/MANG6540 - Integrated Marketing/Final Assessment/Sentiment Analysis/Sentiment_Scores_TokenizedWords.xlsx"
df_combined.to_excel(output_path, index=False)
