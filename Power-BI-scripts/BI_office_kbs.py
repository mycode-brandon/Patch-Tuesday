import pandas as pd
import requests
from bs4 import BeautifulSoup as bs
import re


# For PowerBI
"""
new_kb = "&Text.From(#"This Month's Office KB")&"
"""

# For testing:
new_kb = "5002091"

url = f"https://support.microsoft.com/help/{new_kb}"
req = requests.get(url)
doc = bs(req.text, 'html.parser')
tr = doc.find_all('tr')
product_list = []
num_list = []
for r in tr:
    if r.find('td'):
        td = [x.get_text() for x in r.find_all('td')]
        product = td[0].strip()
        num = td[1].strip()
        num = re.search(r"\(KB(.*)\)", num)
        if num:
            num = num.group(1)
        product_list.append(product)
        num_list.append(num)

d = {'product': product_list, 'num': num_list}
scraped_office_kbs = pd.DataFrame(d)

print(scraped_office_kbs)
