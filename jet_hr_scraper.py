from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver.chrome.options import Options
import logging
import time
from datetime import datetime
import openpyxl
import re
from selenium.common.exceptions import JavascriptException
from openpyxl.styles import Font
import os
from openpyxl import Workbook
import math


# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

headers = [
            "Datetime", "Address URL", "Company Name", "Updated", "Business Establishment",
            "Location", "Fiscal Code", "Legal Form", "Internet Site Name", "NACE Code",
            "Sector", "Legal Representative", "Number of Employees Range", "EM Text",
            "LinkedIn Profile URL", "Presentation", "Competitors", "Region"
]

# Function to get text after label
def get_text_after_label(label):
    element = wait.until(EC.presence_of_element_located((By.XPATH, f'//div[contains(., "{label}")]/following-sibling::div')))
    element_text = element.text
    return element_text





# Set up Chrome options

# Execute the function to click all "Find out more" buttons
try:
    # Initialize the Chrome WebDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    wait = WebDriverWait(driver, 400)
    # Record the start time
    start_time = time.time()
    logging.info("Starting to load the web page.")

    # Open the desired URL directly
    driver.get("https://startup.registroimprese.it/isin/home#")

    driver.set_page_load_timeout(300)


    # Wait for the advanced search link to be clickable
    logging.info("Waiting for the 'advanced search' link to become clickable.")
    advanced_search_link = wait.until(
        EC.element_to_be_clickable((By.ID, "filtroAvanzatoLnk"))
    )

    # Click the advanced search link
    logging.info("'Advanced search' link is clickable. Clicking the link.")
    advanced_search_link.click()

    # Locate the parent div by its id
    # Click the region search link
    logging.info("Finding ID's")
    parent_div = driver.find_element(By.ID, "valueRegione")


    # Clear the inner HTML of the parent div
    children = parent_div.find_elements(By.XPATH, './*')
    get_id = []
    for child in children:
        # Find all grandchildren of each child
        grandchildren = child.find_elements(By.XPATH, './*')
        for grandchild in grandchildren:
            get_id.append(grandchild.get_attribute("id"))
            supergrandchildren = grandchild.find_elements(By.XPATH, './*')
            for supergrandchild in supergrandchildren:
                get_id.append(supergrandchild.get_attribute("id"))

    logging.info("Click Region")
    region_click = wait.until(
        EC.element_to_be_clickable((By.ID, f"{get_id[0]}"))
    )

    children1 = parent_div.find_elements(By.CSS_SELECTOR, 'div[data-value]')


    remove_first_data = 0
    for child1 in children1:
        data_value = child1.get_attribute('data-value')

        print(f'Clicking element with data-value: {data_value}')

        if remove_first_data == 0:
            try:
                try:
                    # Scroll the element into view
                    driver.execute_script("arguments[0].scrollIntoView(true);", child1)
                    time.sleep(0.1)  # Allow some time for the scrolling

                    # Click the element using JavaScript to avoid interactability issues
                    driver.execute_script("arguments[0].click();", child1)

                except ElementNotInteractableException:
                    print(f"Element with data-value {data_value} not interactable.")
            except ElementNotInteractableException:
                    print(f"Element with data-value {data_value} not interactable.")

        else:
            try:
                delete_icon = driver.find_element(By.XPATH, "//i[@class='delete icon']")
                print(f'Clicking element with data-value: {data_value}')
                try:
                    # Scroll the element into view
                    driver.execute_script("arguments[0].scrollIntoView(true);", delete_icon)
                    time.sleep(0.1)  # Allow some time for the scrolling

                    # Click the element using JavaScript to avoid interactability issues
                    driver.execute_script("arguments[0].click();", delete_icon)

                except ElementNotInteractableException:
                    print(f"Element with data-value {data_value} not interactable.")
            except ElementNotInteractableException:
                    print(f"Element with data-value {data_value} not interactable.")
            except Exception as e:
                print(f"An error occurred: {e}")

            try:
                try:
                    # Scroll the element into view
                    driver.execute_script("arguments[0].scrollIntoView(true);", child1)
                    time.sleep(0.1)  # Allow some time for the scrolling

                    # Click the element using JavaScript to avoid interactability issues
                    driver.execute_script("arguments[0].click();", child1)

                except ElementNotInteractableException:
                    print(f"Element with data-value {data_value} not interactable.")
            except ElementNotInteractableException:
                print(f"Element with data-value {data_value} not interactable.")

        # Wait for the checkboxes to be present
        try:
            checkboxes = wait.until(
                EC.presence_of_all_elements_located((By.XPATH, "//input[@type='checkbox']"))
            )
            # Use JavaScript to check each checkbox
            for checkbox in checkboxes[:2]:
                driver.execute_script("arguments[0].checked = true;", checkbox)
        except:
            print("No Check Boxes")


          # Example action: print the text of the span element
        try:
            get_region = wait.until(
                EC.presence_of_element_located((By.XPATH, f"//a[@class='ui label transition visible' and @data-value='{data_value}']/span"))
            )
            # Print the text of the span element to verify
            get_region = get_region.text
        except Exception as e:
            print(f"An error occurred: {e}")

        file_path = rf"D:\Documents\SideProjectFiles\Upwork\jet-hr\{get_region}.xlsx"
        file_exists = os.path.exists(file_path)
        if not file_exists:
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(headers)
        else:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active


        # Wait for the element to be present
        search_button = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "searchBtnVetrina"))
        )

        try:
            # Click the element using JavaScript
            driver.execute_script("arguments[0].click();", search_button)
        except JavascriptException as e:
            print(f"JavaScript click failed: {e}")


        # Wait for the span element containing the text "showing 1 of 341" to be present
        span_element = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[@class='four wide column']//span[contains(text(), 'of')]"))
        )

        # Extract the text from the span element
        span_text = span_element.text

        # Print the extracted text
        parts = span_text.split()
        index_of_of = parts.index("of")
        value_after_of = parts[index_of_of + 1]
        compare_val = f'showing {value_after_of} of {value_after_of}'

        wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class, 'ui button') and contains(@class, 'rounded') and contains(@class, 'black') and @style='background-color: #525252;']")))


        # Find all div elements with the specified class and text
        buttons = driver.find_elements(By.XPATH, "//div[contains(@class, 'ui button') and contains(@class, 'rounded') and contains(@class, 'black') and @style='background-color: #525252;']")
        num_buttons = len(buttons)
        index = 0

        while index < num_buttons:
            # Find the buttons again after each navigation back
            try:
                buttons = driver.find_elements(By.XPATH, "//div[contains(@class, 'ui button') and contains(@class, 'rounded') and contains(@class, 'black') and @style='background-color: #525252;']")
                button_div = buttons[index]
            except:
                driver.back()
                time.sleep(5)
                buttons = driver.find_elements(By.XPATH, "//div[contains(@class, 'ui button') and contains(@class, 'rounded') and contains(@class, 'black') and @style='background-color: #525252;']")
                button_div = buttons[index]


            print(f"Clicking button {index}/{num_buttons}")

            # Scroll to the element
            driver.execute_script("arguments[0].scrollIntoView(true);", button_div)
            driver.execute_script("arguments[0].click();", button_div)

            time.sleep(5 * index)

            # Wait for any potential navigation or page changes (adjust as necessary)
            # Company name
            current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            try:
              company_name = wait.until(EC.presence_of_element_located((By.XPATH, '//span[contains(@id, "companyNameForGA")]'))).text
            except:
              company_name = ""

            try:
                updated_element = wait.until(EC.presence_of_element_located((By.XPATH, '//h2[contains(text(), "BUSINESS REGISTER INFORMATION")]')))
                updated_text = driver.execute_script("return arguments[0].nextSibling.nodeValue;", updated_element).strip()
                updated = updated_text.split("Updated")[1].strip()
            except:
                updated_element = ""
                updated_text = ""
                updated = ""

            business_establishment = get_text_after_label("Business Establishment")
            location = get_text_after_label("Location")
            fiscal_code = get_text_after_label("Fiscal code")
            legal_form = get_text_after_label("Legal form")
            nace_code = get_text_after_label("NACE Code")
            sector = get_text_after_label("Sector")
            number_of_employees_range = wait.until(EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "Number of Employees Range")]/following-sibling::div/span'))).text

            try:
                email_url_elements = driver.find_elements(By.XPATH, "//a[contains(@href, 'mailto:')]")
            except:
                email_url_elements = ""

            try:
                em_el = [email_element for index, email_element in enumerate(email_url_elements) if email_element.text]
                em_url = em_el.get_attribute('href')
                em_text = em_el.text
            except:
                em_url = ""
                em_text = ""

            try:
                legal_representative_element = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'legal representative')]")))
                name = legal_representative_element.text.split(":")[-1].strip() # Extracting the name after the colon and trimming spaces
                pattern = r"by the legal representative (.*?) on"
                match = re.search(pattern, name)
                legal_representative = match.group(1)
            except:
                legal_representative = ""

            try:
                h4_element = driver.find_element(By.XPATH, "//h4[contains(@class, 'header')]//span[text()='PRESENTATION']")
                parent_h4 = h4_element.find_element(By.XPATH, "./ancestor::h4")
                presentation= parent_h4.find_element(By.XPATH, "following-sibling::div[2]").text
            except:
                presentation= ""

            try:
                h4_element_2 = driver.find_element(By.XPATH, "//h4[contains(@class, 'header')]//span[text()='COMPETITORS']")
                parent_h4_2 = h4_element_2.find_element(By.XPATH, "./ancestor::h4")
                competitors= parent_h4_2.find_element(By.XPATH, "following-sibling::div[2]").text
            except:
                competitors= ""

            try:
                linkedin_profile_url = wait.until(EC.presence_of_element_located((By.XPATH, '//a[contains(@href, "https://www.linkedin.com/company")]'))).get_attribute('href')
            except:
                linkedin_profile_url= ""

            try:
                internet_site_name = wait.until(EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "Internet site")]/following-sibling::div/a'))).text
            except:
                internet_site_name= ""

            current_address_url = driver.current_url






            # Define the data to write
            data = [
                        current_datetime, current_address_url, company_name, updated, business_establishment,
                        location, fiscal_code, legal_form, internet_site_name, nace_code, sector, legal_representative,
                        number_of_employees_range, em_text, linkedin_profile_url, presentation, competitors, get_region
                    ]

            next_row = sheet.max_row + 1

            # Write the data to the next empty row
            for col_num, value in enumerate(data, 1):
                cell = sheet.cell(row=next_row, column=col_num, value=value)
                # if value in [linkedin_profile_url, em_text, internet_site_name]:
                #     cell.hyperlink = value
                #     cell.style = "Hyperlink"
                #     cell.font = Font(color="0000FF", underline="single")

            workbook.save(file_path)

            print(f"URL after clicking button {index}: {current_address_url} at {current_datetime}")

            if index == 9:
                driver.back()
                time.sleep(2 * index)
                next_page = wait.until(EC.presence_of_element_located((By.XPATH, "//a[@rel='next' and @style='padding: 0.5em 0.75em' and @title='Go to next page']")))
                driver.execute_script("arguments[0].scrollIntoView(true);", next_page)
                driver.execute_script("arguments[0].click();", next_page)
                time.sleep(2 * index)
                print("current page: " + span_text)
                if span_text == compare_val:
                    remove_first_data = 1
                    index = 0
                    break
                else:
                    index = 0
            else:
                driver.back()
                time.sleep(5 * index)
                index += 1


        end_time = time.time()
        # Calculate the total time taken
        total_time = end_time - start_time
        # Print the total time taken
        logging.info(f"Time taken to access and load the web page: {total_time:.2f} seconds")



except Exception as e:
    logging.error(f"An error occurred: {e}", exc_info=True)

finally:
    # Close the browser
    driver.quit()
    logging.info("Browser closed.")

