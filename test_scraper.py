from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException
import logging
import time
from datetime import datetime
import openpyxl
import traceback
import re
from selenium.webdriver.common.action_chains import ActionChains

# Initialize the WebDriver (make sure to download the appropriate driver for your browser)
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)


# Open the website
driver.get("https://startup.registroimprese.it/isin/home#")

# Wait for the page to load
time.sleep(5)  # Adjust as necessary

# Locate and interact with the region dropdown
region_dropdown = Select(driver.find_element(By.ID, 'region'))  # Adjust ID as necessary

# Iterate over each region
regions = [option.text for option in region_dropdown.options]
data = []

for region in regions:
    region_dropdown.select_by_visible_text(region)
    time.sleep(3)  # Wait for the page to update with new results

    while True:
        # Scrape the data from the updated results
        companies = driver.find_elements(By.CSS_SELECTOR, '.company-class')  # Adjust the selector as necessary

        for company in companies:
            name = company.find_element(By.CSS_SELECTOR, '.name-class').text  # Adjust the selector as necessary
            details = company.find_element(By.CSS_SELECTOR, '.details-class').text  # Adjust the selector as necessary

            # Click "Find out more" to get additional details
            find_out_more_button = company.find_element(By.CSS_SELECTOR, '.find-out-more-class')  # Adjust the selector as necessary
            find_out_more_button.click()
            time.sleep(2)  # Wait for the additional details to load

            # Scrape additional details
            more_details = driver.find_element(By.CSS_SELECTOR, '.more-details-class').text  # Adjust the selector as necessary

            # Close the additional details
            close_button = driver.find_element(By.CSS_SELECTOR, '.close-more-details-class')  # Adjust the selector as necessary
            close_button.click()
            time.sleep(1)  # Wait for the details to close

            # Store all the collected data
            data.append({'region': region, 'name': name, 'details': details, 'more_details': more_details})

        # Check for the presence of a "next page" button and click it if it exists
        try:
            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '.next-page-class'))  # Adjust the selector as necessary
            )
            next_button.click()
            time.sleep(3)  # Wait for the next page to load
        except:
            # If there's no next button, break the loop
            break

driver.quit()

# Process the collected data (e.g., save to a file, print, etc.)
import pandas as pd
df = pd.DataFrame(data)
df.to_csv('startup_data.csv', index=False)

print(df.head())