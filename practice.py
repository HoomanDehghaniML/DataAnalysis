import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pandas.plotting import register_matplotlib_converters

def analyze_csv(file_url):
    # Read the CSV file
    df = pd.read_csv(file_url)

    # Generate some interesting findings
    total_referring_pages = len(df)
    first_referring_page = df['Referring page URL'].iloc[0]
    last_referring_page = df['Referring page URL'].iloc[-1]

    findings = f"""
    Total number of referring pages: {total_referring_pages}
    First referring page: {first_referring_page}
    Last referring page: {last_referring_page}
    """

    # Generate some interesting graphs
    register_matplotlib_converters()

    # Plot of Domain Rating
    plt.figure(figsize=(10, 6))
    sns.histplot(df['Domain rating'], kde=False, bins=30)
    plt.title('Distribution of Domain Ratings')
    plt.xlabel('Domain Rating')
    plt.ylabel('Count')
    plt.savefig('domain_rating.png')

    # Save the findings and graphs in an HTML file
    with open('summary.html', 'w') as f:
        f.write('<html><body>\n')
        f.write('<h1>Summary</h1>\n')
        f.write('<p>' + findings.replace('\n', '<br>') + '</p>\n')
        f.write('<img src="domain_rating.png" alt="Domain Rating">\n')
        f.write('</body></html>')

# Call the function with the URL of your CSV file
analyze_csv('https://pastebin.com/raw/C69DLchq')