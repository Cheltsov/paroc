from bs4 import BeautifulSoup
from urllib.request import urlopen


feed_url = 'https://www.paroc.ru/producty/stroitelnaya-izolyaciya/obschestroitelnaya-teploizolyaciya/paroc-extra'
f = urlopen(feed_url)
feed = f.read()
feed = BeautifulSoup(feed, 'html.parser')

print(feed.find("img", width="180"))
