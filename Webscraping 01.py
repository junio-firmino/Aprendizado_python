from urllib.request import urlopen
from bs4 import BeautifulSoup

page = urlopen('http://www.pythonbrasil.com.br')    # Primeiro webscraping
bs = BeautifulSoup(page.read(), 'html.parser')
list = bs.find('div',{'id':'section-about'})
for i in list:
    print(i)

