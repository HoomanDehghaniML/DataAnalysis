import os
import pandas as pd
from google.cloud import language_v1
from google.cloud.language_v1 import types
import shutil

def NLP(competitors):
    # Set up Google Natural Language API
    os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = r"C:\Users\Hooman Deghani\OneDrive\PC\Desktop\API Keys\natural-language-api-381923-52cfb8b44bd7.json"
    client = language_v1.LanguageServiceClient()

    def get_entities(text):
        document = language_v1.Document(content=text, type_=language_v1.Document.Type.PLAIN_TEXT)
        request = language_v1.AnalyzeEntitiesRequest(document=document)
        response = client.analyze_entities(request=request)
        return [(entity.name, entity.salience) for entity in response.entities]

    def get_categories(text):
        document = language_v1.Document(content=text, type_=language_v1.Document.Type.PLAIN_TEXT)
        request = language_v1.ClassifyTextRequest(document=document)
        response = client.classify_text(request=request)
        return [(category.name, category.confidence) for category in response.categories]

    def get_sentiment(text):
        document = language_v1.Document(content=text, type_=language_v1.Document.Type.PLAIN_TEXT)
        request = language_v1.AnalyzeSentimentRequest(document=document)
        response = client.analyze_sentiment(request=request)
        return response.document_sentiment.score

    def read_text_file(file_path):
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()

    competitor_data = [read_text_file(file_path) for file_path in competitors]

    # Get entity data
    entity_data = []
    for text in competitor_data:
        entity_data.extend(get_entities(text))

    df_competitors = pd.DataFrame(entity_data, columns=['Entity', 'Salience'])
    df_competitors_aggregated = df_competitors.groupby('Entity').mean().reset_index()

    # Sort df_competitors_final by Salience score in descending order
    df_competitors_final = df_competitors_aggregated.sort_values('Salience', ascending=False)

    # Remove rows with Salience < 0.00055
    df_competitors_final = df_competitors_final[df_competitors_final['Salience'] >= 0.00055]

    # Get top 250 entities
    df_competitors_final = df_competitors_final.head(250)

    # Save entity data to xlsx
    df_competitors_final.to_excel("entities.xlsx", index=False, engine='openpyxl')

    # Get category data
    competitor_categories = []
    for text in competitor_data:
        competitor_categories.extend(get_categories(text))

    df_category_competitors = pd.DataFrame(competitor_categories, columns=['Category', 'Confidence score'])
    df_category_competitors_aggregated = df_category_competitors.groupby('Category').mean().reset_index()

    category_html = f"<h1> Category Data </h1> <h2> competitors </h2>{df_category_competitors_aggregated.to_html()}"

    # Get sentiment data
    sentiment_competitors = sum(get_sentiment(text) for text in competitor_data) / len(competitor_data)
    sentiment_html = f"<h1>Average Sentiment Score</h1><p>The average sentiment score of competitors is {sentiment_competitors}.</p>"

    # Save summary to HTML file
    with open('summary.html', 'w') as summary_file:
        summary_file.write("<html><head><style>table, th, td {border: 1px solid black; border-collapse: collapse;} th, td {padding: 10px;} </style></head><body>")
        summary_file.write(category_html)
        if sentiment_html:
            summary_file.write(sentiment_html)
        summary_file.write(f"<h1>Entities</h1><p>A list of the most salient entities to the top performing competitors is attached to this summary.</p>")
        summary_file.write("</body></html>")

    # Create a new folder and save the files
    folder_name = input("Enter the name of the folder: ")
    output_dir = r"G:\Shared drives\sq1 - marketing\seo\Intent Research"
    folder_path = os.path.join(output_dir, folder_name)

    # Create the folder if it doesn't exist
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # Save the files in the new folder
    shutil.move("entities.xlsx", os.path.join(folder_path, "entities.xlsx"))
    shutil.move("summary.html", os.path.join(folder_path, "summary.html"))

    print("Done!")

# User data
competitors_list = [
r"C:\Users\Hooman Deghani\OneDrive\PC\Desktop\Past Projects\2023 Q1\Ai tools for user intent research\flexjobs.txt",
r"C:\Users\Hooman Deghani\OneDrive\PC\Desktop\Past Projects\2023 Q1\Ai tools for user intent research\hubspot.txt",
r"C:\Users\Hooman Deghani\OneDrive\PC\Desktop\Past Projects\2023 Q1\Ai tools for user intent research\investopedia.txt"
]

NLP(competitors_list)
