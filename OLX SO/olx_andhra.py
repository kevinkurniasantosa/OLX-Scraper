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
import urllib
import json

print('Import success\n')
start = time.time()

main_url = 'https://www.olx.in'
url = 'https://www.olx.in/andhra-pradesh_g2001145/for-sale-shops-offices_c1733'
all_product = []

# Selenium
chrome_options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=chrome_options)
driver.get(url)
time.sleep(2)

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
        print('\n=============================\n')
        # Property URL
        try:
            prop_url = prop.find('a', href=True)
            temp = prop_url['href']
            prop_url = main_url + temp
        except Exception as err:
            print('Error parsing URL -> ' + str(err))
        print('Prop URL: ' + str(prop_url))

        driver.quit()

        res = requests.get(prop_url, headers={'User-Agent': 'Mozilla/5.0'})
        soup = BeautifulSoup(res.text, 'html.parser')
     
        # Price
        try:
            price = soup.find('span', class_='_2xKfz').text.strip()
            price = clean_string(price)
        except:
            price = '-'
        print('Price:' + price)

        # State, District, City
        try:
            sdc = soup.find('span', class_='_2FRXm').text.strip()
            x = re.match('(.+), (.+), (.+)', str(sdc))
            # State
            try:
                state = x.group(3)
            except:
                state = '-'
            # District
            try:
                district = x.group(2)
            except:
                district = '-'
            # City
            try:
                city = x.group(1)
            except: 
                city = '-'
        except:
            state = '-'
            district = '-'
            city = '-' 
        print('State: ' + state)
        print('District: ' + district)
        print('City: ' + city)

        # Details
        try:
            details = soup.find('div', class_='_3JPEe')
            # Type
            try:
                prop_type = details.find('span', {'data-aut-id': 'value_type'}).text.strip()
                prop_type = clean_string(prop_type)
            except:
                prop_type = '-'
            # Bedrooms
            try:
                bedrooms = details.find('span', {'data-aut-id': 'value_rooms'}).text.strip()
                bedrooms = clean_string(bedrooms)
            except:
                bedrooms = '-'
            # Bathrooms
            try:
                bathrooms = details.find('span', {'data-aut-id': 'value_bathrooms'}).text.strip()
                bathrooms = clean_string(bathrooms)
            except:
                bathrooms = '-'
            # Furnishing
            try:
                furnishing = details.find('span', {'data-aut-id': 'value_furnished'}).text.strip()
                furnishing = clean_string(furnishing)
            except: 
                furnishing = '-'
            # Status
            try:
                status = details.find('span', {'data-aut-id': 'value_constructionstatus'}).text.strip()
                status = clean_string(status)
            except:
                status = '-'
            # Listed by
            try:
                listed_by = details.find('span', {'data-aut-id': 'value_listed_by'}).text.strip()
                listed_by = clean_string(listed_by)
            except:
                listed_by = '-'
            # Builtup Area
            try:
                builtup_area = details.find('span', {'data-aut-id': 'value_ft'}).text.strip()
                builtup_area = clean_string(builtup_area)
            except:
                builtup_area = '-'
            # Carpet Area
            try:
                carpet_area = details.find('span', {'data-aut-id': 'value_carpetarea'}).text.strip()
                carpet_area = clean_string(carpet_area)
            except:
                carpet_area = '-'    
            # Maintenance
            try: 
                maintenance = details.find('span', {'data-aut-id': 'value_maintenance'}).text.strip()
                maintenance = clean_string(maintenance)
            except:
                maintenance = '-'     
            # Total Floors
            try:
                total_floors = details.find('span', {'data-aut-id': 'value_totalfloors'}).text.strip()
                total_floors = clean_string(total_floors)
            except:
                total_floors = '-'
            # Floor No
            try:
                floor_no = details.find('span', {'data-aut-id': 'value_floorno'}).text.strip()
                floor_no = clean_string(floor_no)
            except:
                floor_no = '-'    
            # Car Parking
            try: 
                car_parking = details.find('span', {'data-aut-id': 'value_carparking'}).text.strip()
                car_parking = clean_string(car_parking)
            except:
                car_parking = '-'     
            # Facing
            try:
                facing = details.find('span', {'data-aut-id': 'value_facing'}).text.strip()
                facing = clean_string(facing)
            except:
                facing = '-'    
            # Project Name
            try: 
                project_name = details.find('span', {'data-aut-id': 'value_projects'}).text.strip()
                project_name = clean_string(project_name)
            except:
                project_name = '-'    
        except Exception as err:
            print('Error in details -> ' + str(err))
        print('Type: ' + prop_type)
        print('Bedrooms: ' + bedrooms)
        print('Bathrooms: ' + bathrooms)
        print('Furnishing: ' + furnishing)
        print('Status: ' + status)
        print('Listed by: ' + listed_by)
        print('Builtup Area: ' + builtup_area)
        print('Carpet Area: ' + carpet_area)
        print('Maintenance: ' + maintenance)
        print('Total Floors: ' + total_floors)
        print('Floor No: ' + floor_no)
        print('Car Parking: ' + car_parking)            
        print('Facing: ' + facing)
        print('Project Name: ' + project_name)
           

        # Description
        try:
            description = soup.find('div', {'data-aut-id': 'itemDescriptionContent'}).text.strip()
            description = clean_string(description)
        except:
            description = '-'
        print('Description: ' + description)

        # AD ID
        try:
            ad_id = soup.find('div', class_='fr4Cy').strong.text.strip()
            x = re.match('AD ID (.+)', str(ad_id))
            ad_id = x.group(1)
            ad_id = clean_string(ad_id)
        except:
            ad_id = '-'
        print('AD ID: ' + ad_id)

        data = {
            'URL': prop_url,
            'Price': price,
            'State': state,
            'District': district,
            'City': city,
            'Type': prop_type,
            'Bedrooms': bedrooms,
            'Bathrooms': bathrooms,
            'Furnishing': furnishing,
            'Status': status,
            'Listed by': listed_by,
            'Builtup Area': builtup_area,
            'Carpet Area': carpet_area,
            'Total Floors': total_floors,
            'Floor No': floor_no,
            'Car Parking': car_parking,
            'Maintenance': maintenance,
            'Facing': facing,
            'Project Name': project_name,
            'Description': description,
            'AD ID': ad_id
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
filename_csv = 'OLX Goa.csv'
with open(filename_csv , 'w', newline='') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    olx_scraping()
            
# Convert from CSV to Excel
filename_excel = 'OLX Goa.xlsx'
merge_all_to_a_book(glob.glob(filename_csv), filename_excel)

end = time.time()
run_time = end - start
run_time_hour = run_time/3600
print('\nScript runs for', round(run_time), 'seconds')
print('Script runs for', round(run_time_hour), 'hour(s)')