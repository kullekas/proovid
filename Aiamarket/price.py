"""Download Aiamarket pricelist and save it to Excel."""

# https://www.youtube.com/watch?v=XQgXKtPSzUI

import urllib.request
import xlsxwriter
from bs4 import BeautifulSoup as soup

print("Laen andmeid!")

my_url = "http://www.aiamarket.ee/best-sales?n=637"

# Avab urli ja tõmbab lehe alla
uClient = urllib.request.urlopen(my_url)
page_html = uClient.read()
uClient.close()

page_soup = soup(page_html, "html.parser")

containers = page_soup.findAll("div", {"class": "product-container"})

contain = containers[0]
container = containers[0]

workbook = xlsxwriter.Workbook('osta.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column(1, 1, 15)
bold = workbook.add_format({'bold': 1})

worksheet.write('A1', 'Toode', bold)
worksheet.write('B1', 'Hind', bold)
worksheet.write("C1", "Koht", bold)

row = 1
col = 0
koht = 1

for container in containers:
    toode = container.div.div.a.img["title"]
    hind1 = container.findAll("span", {"class": "price"})
    hind = hind1[0].text.strip()

    if "Royal" in toode:
        worksheet.write_string(row, col, toode)
        worksheet.write_string(row, col + 1, hind)
        worksheet.write_number(row, col + 2, koht)

        row += 1
        koht += 1

print("Andmed laetud ja salvestatud faili")

workbook.close()
