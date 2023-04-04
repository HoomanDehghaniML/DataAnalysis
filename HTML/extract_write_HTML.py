import requests
from bs4 import BeautifulSoup
import os

MAIN_TAGS = ["main", "article", "content", "post", "entry-content", "main-content"]

def get_main_text(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser")

    for tag in MAIN_TAGS:
        main_content = soup.find(tag)
        if main_content:
            break

    if not main_content:
        return None

    paragraphs = main_content.find_all("p")
    heading_tags = ["h1", "h2", "h3", "h4", "h5", "h6"]
    headings = main_content.find_all(heading_tags)
    lists = main_content.find_all(["ul", "ol"])
    list_items = []
    for lst in lists:
        list_items.extend(lst.find_all("li"))

    # Extract direct text content inside the main tag, excluding any nested tags
    direct_texts = [text for text in main_content.stripped_strings if text not in [t.get_text() for t in main_content.find_all()]]

    text = "\n".join([heading.get_text() for heading in headings] + [p.get_text() for p in paragraphs] + [li.get_text() for li in list_items] + direct_texts)
    return text


def save_to_file(url, content):
    file_name = f"{url.replace('/', '_').replace(':', '_')}.txt"
    with open(file_name, "w", encoding="utf-8") as file:
        file.write(content)

if __name__ == "__main__":
    url = input("Please enter the URL of the article: ")
    main_text = get_main_text(url)
    
    if main_text:
        save_to_file(url, main_text)
        print(f"The main text from '{url}' has been saved to the current directory.")
    else:
        print("Could not find the main content of the page.")
