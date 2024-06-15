import time
import os
import sys
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from datetime import datetime
import csv
import logging

# Configure logging to suppress informational messages
logging.basicConfig(level=logging.WARNING)

load_dotenv()

chrome_options = Options()
chrome_options.add_argument("--headless")  # Run Chrome in headless mode

def main(excel_file=None):
    if excel_file:
        wb = xw.Book(excel_file)
        sheet_name = 'funds_info'
        try:
            sheet = wb.sheets[sheet_name]
            sheet.clear()  # Clear existing sheet
        except:
            sheet = wb.sheets.add(sheet_name)  # Create new sheet if it doesn't exist
    else:
        sheet = None

    html_content = parse_html()
    soup = BeautifulSoup(html_content, 'html.parser')
    funds = get_my_funds(soup)

    if sheet:
        update_excel(sheet, funds)
    else:
        write_csv_file(funds)

def parse_html():
    url = ('https://smartinvestor.sorted.org.nz/kiwisaver-and-managed-funds/'
           '?fundTypes=all-fund-types&managedFundTypes=kiwisaver&sort=growth-assets-asc')

    driver = webdriver.Chrome(options=chrome_options)
    driver.get(url)
    driver.implicitly_wait(20)

    try:
        close_modal_button = driver.find_element(By.CLASS_NAME, "leadinModal-close")
        close_modal_button.click()
    except:
        logging.warning('No modal to close')

    while True:
        try:
            load_more_button = driver.find_element(By.CLASS_NAME, "Pagination__button")
            load_more_button.click()
            time.sleep(5)
        except:
            logging.warning('Finished scraping')
            break

    html_content = driver.page_source
    driver.quit()
    return html_content

def get_my_funds(soup):
    all_funds = soup.find_all('div', class_='FundTile')
    my_funds = []
    seen_funds = set()

    for fund in all_funds:
        fund_values = fund.find_all('div', class_='DoughnutChartWrapper__main-val')
        return_percentage = fund_values[1].text.strip().split('\n')[0].strip().rstrip('%')

        if return_percentage == 'No five-year data available':
            continue

        return_percentage = float(return_percentage)
        fee_percentage = float(fund_values[0].text.strip().split('\n')[0].strip().rstrip('%'))
        provider_name = fund.find('p', class_='FundTile__category').text.strip()
        fund_name = fund.find('h3', class_='FundTile__title').text.strip()
        fund_link = fund.find('a', href=True)['href']

        try:
            fund_category = fund.find_all('span', class_='Tag FundTile__tag')[1].text.strip().split('\n')[-1].strip()
        except IndexError:
            fund_category = 'N/A'

        provider_and_fund = f"{provider_name} {fund_name}"
        
        # Check for duplicates
        fund_key = (provider_name, fund_name)
        if fund_key in seen_funds:
            continue

        seen_funds.add(fund_key)
        my_funds.append([
            provider_and_fund, return_percentage, fund_link, fund_category, fee_percentage, provider_name, fund_name
        ])

    return my_funds

def update_excel(sheet, funds):
    headers = ['Provider and Fund', 'Return % (last 5 years)', 'Fund Link', 'Fund Category', 'Fee %', 'Provider', 'Fund']
    sheet.range('A1').value = headers
    sheet.range('A2').value = funds

def write_csv_file(funds):
    now = datetime.now()
    filename = f"kiwisaverfunds-{now.strftime('%Y-%m-%d-%H-%M-%S')}.csv"

    headers = ['Provider and Fund', 'Return % (last 5 years)', 'Fund Link', 'Fund Category', 'Fee %', 'Provider', 'Fund']
    with open(filename, 'w', encoding='utf-8', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        writer.writerows(funds)

    print(f'Written to CSV file: {filename}')

if __name__ == '__main__':
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        excel_file = None
    main(excel_file)
