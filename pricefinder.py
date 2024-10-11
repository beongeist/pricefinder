import sys
from selenium.webdriver.common.by import By
from selenium.webdriver import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import time
import csv
from sys import exit
import undetected_chromedriver as uc
from fake_useragent import UserAgent
import random as rand

import pandas as pd

# driver_manager = ChromeDriverManager()
# driver_manager.install()



#Options
options = uc.ChromeOptions()
# options.headless=True
# options.add_argument("--headless") 
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.add_argument("--disable-infobars")
options.add_argument("--start-maximized")
options.add_argument("--disable-extensions")
options.add_argument('--window-size=460,853.5')
options.add_argument("--disable-blink-features")
options.add_argument("--disable-blink-features=AutomationControlled")


#Capabilities
capabilities = DesiredCapabilities.CHROME.copy()
capabilities['acceptSslCerts'] = True 
capabilities['acceptInsecureCerts'] = True

# #UserAgent
# ua = UserAgent()
# userAgent = ua.random
# options.add_argument(f'user-agent={userAgent}')

#Driver

chrome_install = ChromeDriverManager().install()
import os
folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")

service = ChromeService(chromedriver_path)

#mainline
driver = webdriver.Chrome(service=service, options=options, desired_capabilities=capabilities)




# driver = webdriver.Chrome()
driver.set_page_load_timeout(10)

#Website Info
loginURL = r'https://www.google.com/search?q=site:pricecharting.com'

#Login
driver.get(loginURL)
time.sleep(rand.uniform(1,3))

import re

def format_search_query(description):
    # Replace special characters with their encoded equivalents
    formatted_query = re.sub(r'[^\w\s]', lambda x: '%%%02x' % ord(x.group()), description)
    return formatted_query

def search(description, grade):
    formatted_description = format_search_query(description)
    search_query = f'site:pricecharting.com {formatted_description}'
    driver.get(f'https://www.google.com/search?q={search_query}')

    # Wait for the search results to load
    WebDriverWait(driver, 80).until(EC.presence_of_element_located((By.XPATH, '//*[@id="rso"]/div[1]/div/div/div/div[1]/div/div/span/a/h3')))

    # Find the first link
    first_link = driver.find_element(By.XPATH, '//*[@id="rso"]/div[1]/div/div/div/div[1]/div/div/span/a/h3')

    # Click on the first link
    time.sleep(rand.uniform(0.1,0.3))
    first_link.click()
    time.sleep(rand.uniform(1,3))

    if grade == 10:
        price_element = driver.find_element_by_xpath("//td[@id='manual_only_price']//span[@class='price js-price']")
        price_value = price_element.text.strip()
        return price_value
    elif grade == 9:
        price_element = driver.find_element_by_xpath("//td[@id='graded_price']//span[@class='price js-price']")
        price_value = price_element.text.strip()
        return price_value
    elif grade == 8:
        price_element = driver.find_element_by_xpath("//td[@id='new_price']//span[@class='price js-price']")
        price_value = price_element.text.strip()
        return price_value
    elif grade == 7:
        price_element = driver.find_element_by_xpath("//td[@id='complete_price']//span[@class='price js-price']")
        price_value = price_element.text.strip()
        return price_value
    else:
        return None
    

import pandas as pd
# Read the Excel file into a pandas DataFrame
df = pd.read_excel("PSA inventory.xlsx")

# Create an empty list to store the updated prices
updated_prices = []

for index, row in df.iterrows():
    # Access the values of the "Description" and "Grade" columns for the current row
    description = row["Description"]
    grade = row["Grade"]

    price = row['APR Value']
    try:
        price = search(description, grade)
        print("{} price for item #{} {}".format(price, index, description))
    except:
        price = 0
        print("Failed to find price for item #{} {}".format(index, description))
    
    # Append the updated price to the list
    updated_prices.append(price)

# Assign the list of updated prices to the "APR Value" column in the DataFrame
df["APR Value"] = updated_prices

# Save the modified DataFrame back to the Excel file
df.to_excel("PSA inventory.xlsx", index=False)



# testSearch = 'Umbreon Vmax Alternate Art'

# search_box = driver.find_element_by_id("game_search_box")

# search_box.send_keys(testSearch)

time.sleep(30)

