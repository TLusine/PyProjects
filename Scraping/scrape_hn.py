import requests
from bs4 import BeautifulSoup
import pprint


def get_custom_hn(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    links = soup.select('.titleline')
    subtext = soup.select('.subtext')

    hn = []
    for idx, item in enumerate(links):
        title = item.getText()
        href = item.a.get('href')
        vote = subtext[idx].select('.score')
        if len(vote):
            points = int(vote[0].getText().replace(' points', ''))
            if points > 99:
                hn.append({'title': title, 'link': href, 'votes': points})
    return hn


def sort_stories_by_votes(hnlist):
    return sorted(hnlist, key=lambda k: k['votes'], reverse=True)


def scrape_multiple_pages(num_pages):
    hn = []
    for page_num in range(1, num_pages + 1):
        url = f'https://news.ycombinator.com/news?p={page_num}'
        hn += get_custom_hn(url)
    return sort_stories_by_votes(hn)


num_pages_to_scrape = 2
result = scrape_multiple_pages(num_pages_to_scrape)
pprint.pprint(result)
