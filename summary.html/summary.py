import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from bs4 import BeautifulSoup
import seaborn as sns

# Load the data
df = pd.read_csv(r"C:\Users\Hooman Deghani\OneDrive\PC\Desktop\internal_all.csv")

output_path = r"G:\Shared drives\sq1 - marketing\seo\Audits\Site Audit\Screaming Frog - Q2 2023"

# Create pie chart for status codes
status_counts = df['Status Code'].value_counts()
fig1, ax1 = plt.subplots()
ax1.pie(status_counts, labels = [f'{k}: {v}' for k, v in zip(status_counts.index, df['Status'].unique())], autopct='%1.1f%%', pctdistance=0.85)
plt.title('Status Codes')
plt.savefig(f'{output_path}/status_codes.png')

# Create pie chart for indexability
index_counts = df['Indexability'].value_counts()
fig2, ax2 = plt.subplots()
ax2.pie(index_counts, labels = index_counts.index, autopct='%1.1f%%', pctdistance=0.85)
plt.title('Indexability')
plt.savefig(f'{output_path}/indexability.png')

# Create pie charts for title and meta description lengths
title_counts = (df['Title 1 Length'] > 60).value_counts()
meta_counts = (df['Meta Description 1 Length'] > 155).value_counts()

fig, axs = plt.subplots(1, 2)
axs[0].pie(title_counts, labels = ['<=60', '>60'], autopct='%1.1f%%')
axs[0].set_title('Title Length')

axs[1].pie(meta_counts, labels = ['<=155', '>155'], autopct='%1.1f%%')
axs[1].set_title('Meta Description Length')

plt.savefig(f'{output_path}/title_and_meta_length.png')

# Create histograms for Flesch Reading Ease and Average Words per Sentence
plt.figure(figsize=(8,8))
sns.histplot(df['Flesch Reading Ease Score'].dropna(), bins=10, kde=True)
plt.title('Flesch Reading Ease Score Distribution')
plt.savefig(f'{output_path}/flesch.png')

plt.figure(figsize=(8,8))
sns.histplot(df['Average Words Per Sentence'].dropna(), bins=10, kde=True)
plt.title('Average Words Per Sentence Distribution')
plt.savefig(f'{output_path}/avg_words.png')

# Create table for addresses and inlinks
df_inlinks = df[['Address', 'Inlinks']].sort_values(by='Inlinks', ascending=False)
df_inlinks.to_html(f'{output_path}/inlinks.html')

# Analyze inlinks distribution
insight = f"5% of pages contain over {int(np.percentile(df['Inlinks'], 95))} links, while 86% pages receive {int(np.percentile(df['Inlinks'], 86))} inlinks or less."

# Create summary.html
soup = BeautifulSoup("", 'html.parser')

# Add titles and images
titles = ['Status Codes', 'Indexability', 'Title Length', 'Meta Description Length', 'Flesch Reading Ease Score Distribution', 'Average Words Per Sentence Distribution', 'Inlinks Insight', 'Inlinks Table']
images = ['status_codes.png', 'indexability.png', 'title_and_meta_length.png', None, 'flesch.png', 'avg_words.png', None, None]
texts = [None, None, None, None, None, None, insight, df_inlinks.to_html()]

for title, image, text in zip(titles, images, texts):
    tag = soup.new_tag('h1')
    tag.string = title
    soup.append(tag)
    if image is not None:
        img_tag = soup.new_tag('img', src=image)
        soup.append(img_tag)
    if text is not None:
        text_tag = BeautifulSoup(text, 'html.parser')
        soup.append(text_tag)

with open(f"{output_path}/summary.html", "w") as file:
    file.write(str(soup.prettify()))