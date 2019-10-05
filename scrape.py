import requests
import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl import load_workbook


if __name__ == '__main__':
    # Fetching currencies
    tgju_currency = 'http://www.tgju.org/currency'
    response = requests.get(tgju_currency)

    soup = BeautifulSoup(response.text, "html.parser")

    dollar_price = soup.select_one(
        '#main > div:nth-child(4) > div > div > div:nth-child(1) > table > tbody > tr:nth-child(1) > td:nth-child(2)'
    ).text
    euro_price = soup.select_one(
        '#main > div:nth-child(4) > div > div > div:nth-child(1) > table > tbody > tr:nth-child(2) > td:nth-child(2)'
    ).text
    dirham_price = soup.select_one(
        '#main > div:nth-child(4) > div > div > div:nth-child(1) > table > tbody > tr:nth-child(4) > td:nth-child(2)'
    ).text
    yuan_price = soup.select_one(
        '#main > div:nth-child(4) > div > div > div:nth-child(1) > table > tbody > tr:nth-child(6) > td:nth-child(2)'
    ).text


    # Fetching petrol prices
    tgju_energy = 'http://www.tgju.org/energy'
    response = requests.get(tgju_energy)

    soup = BeautifulSoup(response.text, "html.parser")

    crude_price = soup.select_one(
        '#main > div:nth-child(4) > div > div > div:nth-child(1) > table > tbody > tr:nth-child(1) > td:nth-child(2)'
    ).text
    brent_price = soup.select_one(
        '#main > div:nth-child(4) > div > div > div:nth-child(1) > table > tbody > tr:nth-child(2) > td:nth-child(2)'
    ).text
    opec_price = soup.select_one(
        '#main > div:nth-child(4) > div > div > div:nth-child(1) > table > tbody > tr:nth-child(3) > td:nth-child(2)'
    ).text
    mazut_price = soup.select_one(
        '#main > div:nth-child(4) > div > div > div:nth-child(1) > table > tbody > tr:nth-child(4) > td:nth-child(2)'
    ).text


    # Fetching global gold prices
    tgju_gold = 'http://www.tgju.org/gold-global'
    response = requests.get(tgju_gold)

    soup = BeautifulSoup(response.text, "html.parser")

    gold_dollars = soup.select_one(
        '#main > div:nth-child(4) > div > div > div:nth-child(1) > table > tbody > tr:nth-child(1) > td:nth-child(2)'
    ).text


    # Writing the scraped data into a spreadsheet

    time_and_date = str(datetime.datetime.now()).split(' ')
    date = time_and_date[0]
    time = time_and_date[1].split('.')[0]

    try:    # Checking if the spreadsheet exists, if so, append the scraped data to it
        wb = load_workbook('records.xlsx')
        ws = wb.active
        ws.append([date, time, dollar_price, euro_price, dirham_price, yuan_price, crude_price, brent_price, opec_price, mazut_price, gold_dollars])
        wb.save('records.xlsx')
    except FileNotFoundError:   # Spreadsheet does not exist in the root folder, creating a new one and appending to it
        wb = Workbook()
        ws = wb.active
        ws.append(['Date', 'Time', 'Dollar Price', 'Euro Price', 'Dirham Price', 'Yuan Price', 'Crude Price', 'Brent Price', 'Opec Price', 'Mazut Price', 'Global Gold Price ($)'])
        ws.append([date, time, dollar_price, euro_price, dirham_price, yuan_price, crude_price, brent_price, opec_price, mazut_price, gold_dollars])
        wb.save('records.xlsx')
