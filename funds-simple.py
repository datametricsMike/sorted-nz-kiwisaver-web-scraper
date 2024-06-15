"""
Generates a CSV file with the following fields for all available KiwiSaver funds:
Provider, Fund Name, Fund Link, Fund Category, Fee %, Return % (last 5 years).
"""
import time
import csv
import os
from datetime import datetime

import selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()

chrome_options = Options()
chrome_options.add_argument("--headless")  # Run Chrome in headless mode

def main():
    """
    Script entry point
    """
    print('Starting')

    html_content = parse_html()
    print('Loaded HTML')

    soup = BeautifulSoup(html_content, 'html.parser')

    print('Starting conversion')
    my_funds = get_my_funds(soup)
    print('Conversion successful')
    write_csv_file(my_funds)
    print('Done')

def parse_html():
    """
    Parse the html with selenium webdriver
    """
    print('Starting to scrape')
    url = ('https://smartinvestor.sorted.org.nz/kiwisaver-and-managed-funds/'
           '?fundTypes=all-fund-types&managedFundTypes=kiwisaver&sort=growth-assets-asc')

    driver = selenium.webdriver.Chrome(options=chrome_options)
    driver.get(url)
    driver.implicitly_wait(20)  # wait 20 seconds max for elements to load

    try:
        close_modal_button = driver.find_element(By.CLASS_NAME, "leadinModal-close")
        close_modal_button.click()
    except selenium.common.exceptions.NoSuchElementException:
        print('No modal to close')

    while True:
        try:
            load_more_button = driver.find_element(By.CLASS_NAME, "Pagination__button")
            load_more_button.click()
            time.sleep(5)
        except selenium.common.exceptions.NoSuchElementException:
            print('Finished scraping')
            break

    print('Starting to parse')
    html_content = driver.page_source
    driver.quit()
    print('Finished parsing')
    return html_content

def get_my_funds(soup):
    """
    Gets a formatted array of all applicable funds
    """
    all_funds = soup.find_all('div', class_='FundTile')
    my_funds = []

    for fund in all_funds:
        current_fund = get_current_fund(fund)

        if current_fund is not None and current_fund not in my_funds:
            my_funds.append(current_fund)

    return my_funds

def get_current_fund(fund):
    """
    Finds values of the current fund and formats them
    """
    fund_values = fund.find_all('div', class_='DoughnutChartWrapper__main-val')
    return_percentage = fund_values[1].text.strip().split('\n')[0].strip().rstrip('%')

    # We skip funds that don't have five-year data
    if return_percentage == 'No five-year data available':
        return None

    return_percentage = float(return_percentage)
    fee_percentage = float(fund_values[0].text.strip().split('\n')[0].strip().rstrip('%'))
    provider_name = fund.find('p', class_='FundTile__category').text.strip()
    fund_name = fund.find('h3', class_='FundTile__title').text.strip()
    fund_link = fund.find('a', href=True)['href']

    try:
        fund_category = fund.find_all(
            'span',
            class_='Tag FundTile__tag'
        )[1].text.strip().split('\n')[-1].strip()
    except IndexError:
        # Some funds don't have an Aggressive, Conservative, etc. tag
        fund_category = 'N/A'

    current_fund = [
        provider_name,
        fund_name,
        fund_link,
        fund_category,
        fee_percentage,
        return_percentage
    ]

    return current_fund

def write_csv_file(my_funds):
    """
    Writes the processed funds to a csv file
    """
    now = datetime.now()
    filename = f"kiwisaverfunds-{now.strftime('%Y-%m-%d-%H:%M:%S')}.csv"
    
    with open(filename, 'w', encoding='utf-8') as file:
        write = csv.writer(file)
        write.writerow(get_headers())
        write.writerows(my_funds)
        print(f'Written to CSV file: {filename}')

def get_headers():
    """
    Returns the CSV column headers
    """
    return [
        'Provider',
        'Fund Name',
        'Fund Link',
        'Fund Category',
        'Fee %',
        'Return % (last 5 years)'
    ]

if __name__ == '__main__':
    main()
