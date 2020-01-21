'''
from bs4 import BeautifulSoup
from urllib.request import urlopen


feed_url = 'https://www.paroc.ru/producty/stroitelnaya-izolyaciya/obschestroitelnaya-teploizolyaciya/paroc-extra'
f = urlopen(feed_url)
feed = f.read()
feed = BeautifulSoup(feed, 'html.parser')
'''



def hash_word(word):
    import hashlib
    word = word.encode('utf-8')
    h = hashlib.sha1(word)
    print(h.hexdigest())


hash_word("привет")
