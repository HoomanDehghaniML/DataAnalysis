import requests
from bs4 import BeautifulSoup
import urllib.parse
import csv

def get_internal_links(url):
    parsed_url = urllib.parse.urlparse(url)
    base_url = parsed_url.scheme + "://" + parsed_url.hostname
    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')
    
    tags = ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'div', 'span', 'li']
    
    internal_links = []
    
    for tag in tags:
        elements = soup.find_all(tag)
        for el in elements:
            for link in el.find_all('a', href=True):
                link_url = link['href']
                parsed_link = urllib.parse.urlparse(link_url)
                if parsed_link.hostname == parsed_url.hostname or parsed_link.hostname is None:
                    full_url = urllib.parse.urljoin(base_url, link_url)
                    if full_url not in internal_links:
                        internal_links.append(full_url)
    
    return internal_links

def save_links_to_csv(links, filename):
    with open(filename, 'w', newline='') as f:
        writer = csv.writer(f)
        for link in links:
            writer.writerow([link])

if __name__ == "__main__":
    url = input("Enter the URL: ")
    links = get_internal_links(url)
    save_links_to_csv(links, 'internal_links.csv')
    print("Links saved to internal_links.csv")
