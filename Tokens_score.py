import pandas as pd

# Replace 'your_file_path' with the actual path to your Excel file
file_path = r"E:\Study Abroad\Semester 01\MANG6540 - Integrated Marketing\Final Assessment\Sentiment Analysis\Sentiment_Scores_TokenizedWords.xlsx"

# Load the Excel file into a DataFrame
df = pd.read_excel(file_path)

# Convert string representation of lists to actual lists
df['tokens'] = df['tokens'].apply(ast.literal_eval)
df['sentiment_scores'] = df['sentiment_scores'].apply(ast.literal_eval)

# Create a new DataFrame with each token and its corresponding sentiment score
data = []
for tokens, sentiment_scores in zip(df['tokens'], df['sentiment_scores']):
    for token, score in zip(tokens, sentiment_scores):
        data.append({'token': token, 'sentiment_score': score})

# Convert the list of dictionaries to a new DataFrame
df_result = pd.DataFrame(data)

# Save the DataFrame to a new Excel file
output_file_path = r"E:\Study Abroad\Semester 01\MANG6540 - Integrated Marketing\Final Assessment\Sentiment Analysis\Sentiment_Scores_TokenizedWords_Positional.xlsx"
df_result.to_excel(output_file_path, index=False)

# Display the resulting DataFrame to verify the data has been processed correctly
print(df_result.head())
