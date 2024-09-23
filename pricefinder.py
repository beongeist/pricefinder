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
# WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "siteHeader__logo")))
# driver.quit()










# #Open Giveaway Page
# driver.get(giveawayURL)
# time.sleep(0.1)
# print("Giveaway Page Screenshoted for Diagnosis Use")
# driver.save_screenshot("Screenshot.png")
# driver.execute_script("window.scrollTo(0, document.body.scrollHeight);") #Check if extraneous

# #Expanding Giveaway Page to Reveal all Entries
# print("Identifying Giveaways . . .")
# cssLoadmore = '#__next > div.PageFrame.PageFrame--siteHeaderBanner > main > div.GiveawayIndexPage__content > div.GiveawayIndexPage__content--main > div.GiveawayIndexPage__listSection > div.Divider.Divider--contents > div > button'
# xpathLoadmore = '//*[@id="__next"]/div[2]/main/div[3]/div[1]/div[2]/div[3]/div/button'
# expected = 0
# print("Expanding Giveaway Page . . . \n")
# expandCount = 0
# while expandCount < 20000:
#     try:
#         loadMoreGiveaways = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpathLoadmore)))
#         driver.execute_script("arguments[0].click();", loadMoreGiveaways)
#         print(expandCount)
#         expandCount += 1
#     except:
#         print("Page Expanded!\n")
#         break



# #Giveaway Information
# giveawayButtons = driver.find_elements(By.XPATH, '//*[@id="__next"]/div[2]/main/div[3]/div[1]/div[2]/div[1]/article/div[2]/div[3]/div[2]/a')
# print("There are {} giveaways available!".format(len(giveawayButtons)))

# giveawayTitle = driver.find_elements(By.XPATH, '//*[@id="__next"]/div[2]/main/div[3]/div[1]/div[2]/div[1]/article/div[2]/div[1]/h3/strong/a')
# giveawayTitle = [titleElement.text for titleElement in giveawayTitle]

# giveawayAuthor = driver.find_elements(By.XPATH, '//*[@id="__next"]/div[2]/main/div[3]/div[1]/div[2]/div[1]/article/div[2]/div[2]/h3/div/span[1]/a/span[1]')
# giveawayAuthor = [authorElement.text for authorElement in giveawayAuthor]
# #driver.save_screenshot("Screenshot.png")


# """ alignmentChecker = (len(giveawayButtons), len(giveawayTitle), len(giveawayAuthor))
# print(alignmentChecker) """

# errors = 0
# entries = 0

# #Entering Giveaways
# for entry, title, author in tuple(zip(giveawayButtons, giveawayTitle, giveawayAuthor)):
#     time.sleep(0.1)
#     driver.execute_script("arguments[0].click();", entry)
#     mainWindow= driver.window_handles[0]
#     infoWindow = driver.window_handles[1]
#     driver.switch_to.window(infoWindow)
#     print("Entering {} written by {}".format(title[:18].center(20), author[:15].center(17)), end=" ")
#     try:
#         time.sleep(0.1)
#         address_button = WebDriverWait(driver,1).until(EC.element_to_be_clickable((By.CLASS_NAME,'addressLink')))
#         driver.execute_script("arguments[0].click();", address_button)
#         time.sleep(0.1)
#         checkbox = WebDriverWait(driver,1).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="termsCheckBox"]')))
#         driver.execute_script("arguments[0].click();", checkbox)
#         time.sleep(0.1)
#         submit_button = WebDriverWait(driver,1).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="giveawaySubmitButton"]')))
#         driver.execute_script("arguments[0].click();", submit_button)
#     except:
#         print("FAILED")
#         with open("GiveawayEntryLog.csv", "a") as logCSV:
#             wr = csv.writer(logCSV)
#             wr.writerow([title, "Failed"])
#             logCSV.close()
#         errors += 1
#     else:
#         print("ENTERED")
#         with open("GiveawayEntryLog.csv", "a", encoding="utf-8") as logCSV:
#             wr = csv.writer(logCSV)
#             wr.writerow([title, "Entered"])
#             logCSV.close()
#         entries += 1
#     driver.close()
#     driver.switch_to.window(mainWindow)

# print("COMPLETE \n{} total scanned \n{} total entered \n{} total failed \n{} total expected".format(errors+entries, entries, errors, len(giveawayButtons)))
# driver.quit()
# # input("Press enter to quit")
