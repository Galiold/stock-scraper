import requests
import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl import load_workbook

SPREADSHEET_NAME = 'test3.xlsx'


def fetch_price_from(url):
    soup = BeautifulSoup(requests.get(url).text, "html.parser")
    return(soup.select_one(
        '#main > div.profile-container.container > div:nth-child(2) > ul > li:nth-child(1) > span'
    ).text)


if __name__ == '__main__':
    # Fetching currencies
    dollar_price = fetch_price_from('https://www.tgju.org/chart/price_dollar_rl')
    euro_price = fetch_price_from('https://www.tgju.org/chart/price_eur')
    dirham_price = fetch_price_from('https://www.tgju.org/chart/price_aed')
    yuan_price = fetch_price_from('https://www.tgju.org/chart/price_cny')

    # Fetching petrol prices
    crude_price = fetch_price_from('https://www.tgju.org/chart/energy-crude-oil')
    brent_price = fetch_price_from('https://www.tgju.org/chart/energy-brent-oil')
    opec_price = fetch_price_from('https://www.tgju.org/chart/oil_opec')
    mazut_price = fetch_price_from('https://www.tgju.org/chart/energy-heating-oil')

    # Fetching global gold prices
    gold_dollars = fetch_price_from('https://www.tgju.org/chart/ons')

    # Writing the scraped data into a spreadsheet
    scraped_data = [datetime.datetime.now().date(), dollar_price, euro_price, dirham_price, yuan_price, crude_price, brent_price, opec_price, mazut_price, gold_dollars]

    # Checking if the spreadsheet exists, if so, append the scraped data to it
    try:
        wb = load_workbook(SPREADSHEET_NAME)
        ws = wb.active

        ws.insert_rows(2)
        for i in range(len(scraped_data)):
            ws.cell(row=2, column=i+1).value = scraped_data[i]

        wb.save(SPREADSHEET_NAME)
        print('Data added to ' + SPREADSHEET_NAME)

    # Spreadsheet does not exist in the root folder, creating a new one and appending to it
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.append(['Datetime', 'Dollar Price', 'Euro Price', 'Dirham Price', 'Yuan Price', 'Crude Price', 'Brent Price', 'Opec Price', 'Mazut Price', 'Global Gold Price ($)'])
        ws.append(scraped_data)
        wb.save(SPREADSHEET_NAME)
        print('Data written in ' + SPREADSHEET_NAME)
