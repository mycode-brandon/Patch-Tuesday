import pandas as pd
import aiohttp
import asyncio
import re
import requests
from bs4 import BeautifulSoup as bs

# For testing
new_kb = "5002087"
start = "2023-03-15"
end = "2023-04-12"
tuesday = "April 11, 2023"


# For PowerBI:
"""
new_kb = "&Text.From(#"This Month's Office KB")&"
start = "&Text.From(#"Report Start")&"
end = "&#"Report End"&"
tuesday = "&#"Patch Tuesday String"&"
"""

deployment_url = f"https://api.msrc.microsoft.com/sug/v2.0/en-US/deployment/?%24orderBy=product+desc&%24filter=%28releaseDate+gt+{start}T00%3A00%3A00-05%3A00%29+and+%28releaseDate+lt+{end}T23%3A59%3A59-05%3A00%29"


def get_kb_data(url):
    req = requests.get(url)
    json = req.json()
    full_kb_data = pd.DataFrame(json['value'])  # creates dataframe from full json values
    return full_kb_data


async def gather_title(kb: str, session: aiohttp.ClientSession):
    url = f"https://support.microsoft.com/help/{kb}"
    async with session.get(url) as article:
        html = await article.text()
        doc = bs(html, 'html.parser')
        title = doc.find('title')
        if title:
            date_title = re.search(r"([A-Z][a-z]+ \d\d?, ?\d\d\d\d)", title.text)
            if date_title:
                date_title = date_title.group(1)
            else:
                date_title = title.text
        date_title = date_title.strip()
        titles_dict = {'articleName': kb}
        if tuesday not in date_title:
            titles_dict['Title'] = date_title.strip()
        else:
            titles_dict['Title'] = ""
        return titles_dict


async def gather_titles(kb_list: list[str]):
    titles = []
    connector = aiohttp.TCPConnector(force_close=True, limit=150)
    async with aiohttp.ClientSession(connector=connector) as session:
        for kb in kb_list:
            titles.append(asyncio.create_task(gather_title(kb, session)))
        await asyncio.gather(*titles)
    dict_list = []
    for item in titles:
        dict_list.append(item.result())
    titles_df = pd.DataFrame(dict_list)
    return titles_df


kb_df = get_kb_data(deployment_url)

kb_nums = kb_df['articleName'].tolist()
kb_nums = list(set(kb_nums))
kb_nums = [x for x in kb_nums if str(x) != 'nan']
for kb in kb_nums:
    reg = re.search(r"^\d+$", kb)
    if not reg:
        kb_nums.remove(kb)
kb_nums = kb_nums


re_released_df = asyncio.run(gather_titles(kb_nums))

print(kb_df)
print(re_released_df)
