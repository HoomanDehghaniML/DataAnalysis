import os
import pandas as pd
from google.cloud import language_v1
from google.cloud.language_v1 import types
import zipfile

def NLP(competitors, sq1):
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
    sq1_data = read_text_file(sq1)

    # Get entity data
    entity_data = []
    for text in competitor_data:
        entity_data.extend(get_entities(text))

    df_competitors = pd.DataFrame(entity_data, columns=['Entity', 'Salience'])
    df_competitors_aggregated = df_competitors.groupby('Entity').mean().reset_index()

    df_sq1 = pd.DataFrame(get_entities(sq1_data), columns=['Entity', 'Salience'])

    common_entities = df_competitors_aggregated[df_competitors_aggregated['Entity'].isin(df_sq1['Entity'])]
    df_competitors_final = df_competitors_aggregated[~df_competitors_aggregated['Entity'].isin(common_entities['Entity'])]

    # Sort df_competitors_final by Salience score in descending order
    df_competitors_final = df_competitors_final.sort_values('Salience', ascending=False)

    # Remove rows with Salience < 0.00055
    df_competitors_final = df_competitors_final[df_competitors_final['Salience'] >= 0.00055]

    # Save entity data to CSV
    project = input("Enter the name of the project: ")
    df_competitors_final.to_csv(f"{project}.csv", index=False)

    # Get category data
    competitor_categories = []
    for text in competitor_data:
        competitor_categories.extend(get_categories(text))

    df_category_competitors = pd.DataFrame(competitor_categories, columns=['Category', 'Confidence score'])
    df_category_competitors_aggregated = df_category_competitors.groupby('Category').mean().reset_index()
    df_category_sq1 = pd.DataFrame(get_categories(sq1_data), columns=['Category', 'Confidence score'])

    category_html = f"<h1> Category Data </h1> <h2> competitors </h2>{ df_category_competitors_aggregated.to_html()} <h2> sq1 </h2>{df_category_sq1.to_html()} <p> If the top three categories from both tables are different, then the article needs an <strong>edit</strong>."

    
    # Get sentiment data
    sentiment_competitors = sum(get_sentiment(text) for text in competitor_data) / len(competitor_data)
    sentiment_sq1 = get_sentiment(sq1_data)
    sentiment_difference = sentiment_competitors - sentiment_sq1

    sentiment_html = f"<h1>Sentiment Difference</h1><p>Sentiment difference is {sentiment_difference}.\n"
    if abs(sentiment_difference) > 0.1:
        sentiment_html += "<p>We need to <strong>edit</strong> the article.</p>"

    # Save summary to HTML file
    with open('summary.html', 'w') as summary_file:
        summary_file.write("<html><head><style>table, th, td {border: 1px solid black; border-collapse: collapse;} th, td {padding: 10px;} </style></head><body>")
        summary_file.write(category_html)
        if sentiment_html:
            summary_file.write(sentiment_html)
        summary_file.write(f"<h1>Edit</h1><p>Go through the csv file and add any appropriate entities to the sq1 article. You can check the competitor articles to find where they have used the entities. Feel free to add/remove sentences to accommodate the new entities.</p>")
        summary_file.write("</body></html>")

    # Compress files into a zip file
    zipfile_name = input("Enter the name of the zipfile: ")
    output_dir = r"G:\Shared drives\sq1 - marketing\seo\Intent Research"
    with zipfile.ZipFile(os.path.join(output_dir, f"{zipfile_name}.zip"), 'w') as archive:
        archive.write(f"{project}.csv")
        archive.write('summary.html')

    # Remove the CSV and HTML files
    os.remove(f"{project}.csv")
    os.remove('summary.html')

    print("Done!")

# User input
competitors_list = [
    r"C:\Users\Hooman Deghani\Python\Data Analysis\HTML\https___www.investopedia.com_updates_first-time-home-buyer_.txt",
    r"C:\Users\Hooman Deghani\Python\Data Analysis\HTML\https___www.nerdwallet.com_article_mortgages_tips-for-first-time-home-buyers.txt",
    r"C:\Users\Hooman Deghani\Python\Data Analysis\HTML\https___www.rocketmortgage.com_learn_first-time-home-buyer-tips.txt"
]
input_sq1 =  r"C:\Users\Hooman Deghani\Python\Data Analysis\HTML\https___www.squareone.ca_resource-centres_home-buying-selling-moving_buying-a-home-for-the-first-time.txt"

# Run the NLP function
NLP(competitors_list, input_sq1)
