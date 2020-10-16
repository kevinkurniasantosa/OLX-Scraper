import random
import string
import datetime
import requests
import re
from bs4 import BeautifulSoup
import time
import csv
import unicodedata
import glob
from pyexcel.cookbook import merge_all_to_a_book
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options

print('Import success\n')
start = time.time()

url = 'https://www.olx.in/andaman-nicobar-islands_g2007598/for-sale-houses-apartments_c1725'
all_product = []

# Selenium
chrome_options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=chrome_options)
driver.get(url)

def clean_string(x):
    try:
        x = x.replace('\"','\'\'').replace('\r',' ').replace('\n',' ').replace(';', ' ')
        x = unicodedata.normalize('NFKD', x).encode('ascii', 'ignore')
        x = x.decode('ascii')
    except:
        x = '?'
    return x

def olx_scraping():
    print('Main URL -> ' + url)

    try:
        while True:
            print('Click load more button')
            load_button = driver.find_element_by_xpath("//button[@class='rui-3sH3b rui-23TLR rui-1zK8h']")
            load_button.click()
            time.sleep(3)
    except:
        pass

    # driver.quit()

    soup_url = BeautifulSoup(driver.page_source, 'html.parser')
    list_property = soup_url.find_all('li', class_='EIR5N')
    print('Number of properties: ' + str(len(list_property)))

    for prop in list_property:
        try:
            # Property URL
            prop_url = prop.find('a', href=True)
            temp = prop_url['href']
            prop_url = main_url + temp
        except Exception as err:
            print('Error parsing URL -> ' + str(err))
        print('Prop URL: ' + prop_url)

        res = requests.get(prop_url, headers={'User-Agent': 'Mozilla/5.0'})
        soup = BeautifulSoup(res.text, 'html.parser')

        # Price
        try:
            price = soup.find('span', class_='_2xKfz').text.strip()
        except:
            price = '-'
        print('Price: ' + price)

        
        data = {
            'URL': url_url,
            'Price-in-Letter': price_letter,
            'Title': title,
            'Status': status,
            'Numeric-Price': numeric_price,
            'Old-ID': old_id,
            'Type': category_type,
            'Area-full': area_full,
            'Bedrooms': bedrooms,
            'Bathrooms': bathrooms,
            'Feature-1': feature_1,
            'Feature-2': feature_2,
            'Feature-3': feature_3,
            'State': state,
            'City': city,
            'District': district,
            'Contact-Name': contact_name,
            'Contact-Phone': contact_phone
        }
        
        all_product.append(data)
        writer.writerow(data)

time.sleep(1)
print('\n==================')
print(all_product)

# Initialize fieldnames
fieldnames=[
    'URL',
    'Price',
    'State',
    'District',
    'City',
    'Type',
    'Bedrooms',
    'Bathrooms',
    'Furnishing',
    'Status',
    'Listed by',
    'Builtup Area',
    'Carpet Area',
    'Total Floors',
    'Floor No',
    'Car Parking',
    'Maintenance',
    'Facing',
    'Project Name',
    'Description',
    'AD ID'
]

# Generate CSV
filename_csv = 'OLX Andaman.csv'
with open(filename_csv , 'w', newline='') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    olx_scraping()
            
# Convert from CSV to Excel
filename_excel = 'OLX Andaman.xlsx'
merge_all_to_a_book(glob.glob(filename_csv), filename_excel)

end = time.time()
run_time = end - start
run_time_hour = run_time/3600
print('\nScript runs for', round(run_time), 'seconds')
print('Script runs for', round(run_time_hour), 'hour(s)')