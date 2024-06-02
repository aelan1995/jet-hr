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
from selenium.common.exceptions import JavascriptException
from openpyxl.styles import Font

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# Function to get text after label
def get_text_after_label(label):
    print(f"Looking for label: {label}")  # Debugging line
    # try:
    #     # Use a more general XPath
    element = wait.until(EC.presence_of_element_located((By.XPATH, f'//div[contains(., "{label}")]/following-sibling::div')))
    element_text = element.text
        #print(f"Found text: {element_text}")  # Debugging line
    return element_text
    # except Exception as e:
    #     #print(f"Failed to find label {label}: {e}")
    #     #print("Page HTML around the label:")
    #     try:
    #         surrounding_html = driver.find_element(By.XPATH, f'//div[contains(., "{label}")]').get_attribute('outerHTML')
    #         #print(surrounding_html)
    #     except:
    #         print("Couldn't find the surrounding HTML for the label.")
    #     return None


def get_and_save_data_to_excel(buttons, num_bottons):
    for index in range(len(buttons)):
        # Find the buttons again after each navigation back
        buttons = driver.find_elements(By.XPATH, "//div[contains(@class, 'ui button') and contains(@class, 'rounded') and contains(@class, 'black') and @style='background-color: #525252;']")
        button_div = buttons[index]
        print(f"Clicking button {index}/{num_bottons}")

        # Scroll to the element
        driver.execute_script("arguments[0].scrollIntoView(true);", button_div)
        driver.execute_script("arguments[0].click();", button_div)

        # Wait for any potential navigation or page changes (adjust as necessary)
        time.sleep(60)

        # Company name
        current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        current_address_url = driver.current_url
        company_name = wait.until(EC.presence_of_element_located((By.XPATH, '//span[contains(@id, "companyNameForGA")]'))).text

        # Updated date
        updated_element = wait.until(EC.presence_of_element_located((By.XPATH, '//h2[contains(text(), "BUSINESS REGISTER INFORMATION")]')))
        updated_text = driver.execute_script("return arguments[0].nextSibling.nodeValue;", updated_element).strip()
        updated = updated_text.split("Updated")[1].strip()

        business_establishment = get_text_after_label("Business Establishment")
        location = get_text_after_label("Location")
        fiscal_code = get_text_after_label("Fiscal code")
        legal_form = get_text_after_label("Legal form")
        internet_site_name = wait.until(EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "Internet site")]/following-sibling::div/a'))).text
        internet_site = 'https://'+internet_site_name
        nace_code = get_text_after_label("NACE Code")
        sector = get_text_after_label("Sector")

        number_of_employees_range = wait.until(EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "Number of Employees Range")]/following-sibling::div/span'))).text
        email = wait.until(EC.presence_of_element_located((By.XPATH, '//a[contains(@href, "mailto:")]'))).get_attribute('href')


        try:
            legal_representative_element = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'legal representative')]")))
            name = legal_representative_element.text.split(":")[-1].strip()  # Extracting the name after the colon and trimming spaces
            pattern = r"by the legal representative (.*?) on"
            match = re.search(pattern, name)
            legal_representative = match.group(1)
            h4_element = driver.find_element(By.XPATH, "//h4[contains(@class, 'header')]//span[text()='PRESENTATION']")
            parent_h4 = h4_element.find_element(By.XPATH, "./ancestor::h4")
            presentation= parent_h4.find_element(By.XPATH, "following-sibling::div[2]").text
            h4_element_2 = driver.find_element(By.XPATH, "//h4[contains(@class, 'header')]//span[text()='COMPETITORS']")
            parent_h4_2 = h4_element_2.find_element(By.XPATH, "./ancestor::h4")
            competitors= parent_h4_2.find_element(By.XPATH, "following-sibling::div[2]").text
            linkedin_profile_url = wait.until(EC.presence_of_element_located((By.XPATH, '//a[contains(@href, "https://www.linkedin.com/company/")]'))).get_attribute('href')
        except:
            # Assign a blank string if the section is not found
            presentation = ""
            competitors = ""
            legal_representative = ""
            linkedin_profile_url = ""

        file_path = "sample-output.xlsx"
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        next_row = sheet.max_row + 1
        # Define the data to write
        data = [
                    current_datetime, current_address_url, company_name, updated, business_establishment,
                    location, fiscal_code, legal_form, internet_site, nace_code, sector, legal_representative,
                    number_of_employees_range, email, linkedin_profile_url, presentation, competitors
                   ]

        # Write the data to the next empty row
        for col_num, value in enumerate(data, 1):
            sheet.cell(row=next_row, column=col_num, value=value)
            # Save the workbook
            workbook.save(file_path)

        print(f"URL after clicking button {index}: {current_address_url} at {current_datetime}")
        if index == 10:
           return True

    return False

# Execute the function to click all "Find out more" buttons
try:
    # Initialize the Chrome WebDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    # Record the start time
    start_time = time.time()
    logging.info("Starting to load the web page.")

    # Open the desired URL directly
    driver.get("https://startup.registroimprese.it/isin/home#")

    driver.set_page_load_timeout(300)


    # Wait for the advanced search link to be clickable
    logging.info("Waiting for the 'advanced search' link to become clickable.")
    advanced_search_link = WebDriverWait(driver, 100).until(
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
    region_click = WebDriverWait(driver, 100).until(
        EC.element_to_be_clickable((By.ID, f"{get_id[0]}"))
    )

    children1 = parent_div.find_elements(By.CSS_SELECTOR, 'div[data-value]')

    for child1 in children1:
        data_value = child1.get_attribute('data-value')
        print(f'Clicking element with data-value: {data_value}')

        try:
            # Scroll the element into view
            driver.execute_script("arguments[0].scrollIntoView(true);", child1)
            time.sleep(0.1)  # Allow some time for the scrolling

            # Click the element using JavaScript to avoid interactability issues
            driver.execute_script("arguments[0].click();", child1)

        except ElementNotInteractableException:
            print(f"Element with data-value {data_value} not interactable.")

    # Wait for the checkboxes to be present
    try:
        checkboxes = WebDriverWait(driver, 100).until(
            EC.presence_of_all_elements_located((By.XPATH, "//input[@type='checkbox']"))
        )
        # Use JavaScript to check each checkbox
        for checkbox in checkboxes[:2]:
            driver.execute_script("arguments[0].checked = true;", checkbox)
    except:
        print("No Check Boxes")


    # Wait for the element to be present
    search_button = WebDriverWait(driver, 100).until(
        EC.presence_of_element_located((By.CLASS_NAME, "searchBtnVetrina"))
    )

    try:
        # Click the element using JavaScript
        driver.execute_script("arguments[0].click();", search_button)
    except JavascriptException as e:
        print(f"JavaScript click failed: {e}")


    wait = WebDriverWait(driver, 400)
    wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class, 'ui button') and contains(@class, 'rounded') and contains(@class, 'black') and @style='background-color: #525252;']")))

    # Find all div elements with the specified class and text
    buttons = driver.find_elements(By.XPATH, "//div[contains(@class, 'ui button') and contains(@class, 'rounded') and contains(@class, 'black') and @style='background-color: #525252;']")
    num_buttons = len(buttons)
    index = 0

    while index < num_buttons:
        # Find the buttons again after each navigation back
        buttons = driver.find_elements(By.XPATH, "//div[contains(@class, 'ui button') and contains(@class, 'rounded') and contains(@class, 'black') and @style='background-color: #525252;']")
        button_div = buttons[index]
        print(f"Clicking button {index}/{num_buttons}")

        # Scroll to the element
        driver.execute_script("arguments[0].scrollIntoView(true);", button_div)
        driver.execute_script("arguments[0].click();", button_div)

        # Wait for any potential navigation or page changes (adjust as necessary)
        time.sleep(10 * index)

        # Company name
        current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        current_address_url = driver.current_url
        company_name = wait.until(EC.presence_of_element_located((By.XPATH, '//span[contains(@id, "companyNameForGA")]'))).text

        # Updated date
        updated_element = wait.until(EC.presence_of_element_located((By.XPATH, '//h2[contains(text(), "BUSINESS REGISTER INFORMATION")]')))
        updated_text = driver.execute_script("return arguments[0].nextSibling.nodeValue;", updated_element).strip()
        updated = updated_text.split("Updated")[1].strip()

        business_establishment = get_text_after_label("Business Establishment")
        location = get_text_after_label("Location")
        fiscal_code = get_text_after_label("Fiscal code")
        legal_form = get_text_after_label("Legal form")
        internet_site_name = wait.until(EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "Internet site")]/following-sibling::div/a'))).text
        internet_site = 'https://'+internet_site_name
        nace_code = get_text_after_label("NACE Code")
        sector = get_text_after_label("Sector")

        number_of_employees_range = wait.until(EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "Number of Employees Range")]/following-sibling::div/span'))).text



        try:
            email = wait.until(EC.presence_of_element_located((By.XPATH, '//a[contains(@href, "mailto:")]')))
            email_href = email.get_attribute('href')
            email_text = email.text
            legal_representative_element = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'legal representative')]")))
            name = legal_representative_element.text.split(":")[-1].strip()  # Extracting the name after the colon and trimming spaces
            pattern = r"by the legal representative (.*?) on"
            match = re.search(pattern, name)
            legal_representative = match.group(1)
            h4_element = driver.find_element(By.XPATH, "//h4[contains(@class, 'header')]//span[text()='PRESENTATION']")
            parent_h4 = h4_element.find_element(By.XPATH, "./ancestor::h4")
            presentation= parent_h4.find_element(By.XPATH, "following-sibling::div[2]").text
            h4_element_2 = driver.find_element(By.XPATH, "//h4[contains(@class, 'header')]//span[text()='COMPETITORS']")
            parent_h4_2 = h4_element_2.find_element(By.XPATH, "./ancestor::h4")
            competitors= parent_h4_2.find_element(By.XPATH, "following-sibling::div[2]").text
            linkedin_profile_url = wait.until(EC.presence_of_element_located((By.XPATH, '//a[contains(@href, "https://www.linkedin.com/")]')))
            linkedin_href = linkedin_profile_url.get_attribute('href')
            linkedin_text = linkedin_profile_url.text

        except:
            # Assign a blank string if the section is not found
            email_href = ""
            email_text = ""
            presentation = ""
            competitors = ""
            legal_representative = ""
            linkedin_profile_url = ""
            linkedin_href = ""
            linkedin_text = ""



        file_path = "sample-output.xlsx"
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        next_row = sheet.max_row + 1
        # Define the data to write
        data = [
                    current_datetime, current_address_url, company_name, updated, business_establishment,
                    location, fiscal_code, legal_form, internet_site, nace_code, sector, legal_representative,
                    number_of_employees_range, email, linkedin_profile_url, presentation, competitors
                   ]

        # Write the data to the next empty row
        converted_data = []
        for value in data:
            if isinstance(value, (int, float, str, datetime, type(None))):
                converted_data.append(value)
            else:
                converted_data.append(str(value))
        for col_num, value in enumerate(converted_data, 1):
            cell = sheet.cell(row=next_row, column=col_num, value=value)
            # Save the workbook
            if col_num == 13:  # Assuming the email column is the 13th column
                cell.hyperlink = f"mailto:{email_href}"
                cell.value = email_text
                cell.font = Font(color="0000FF", underline="single")
            if col_num == 14:  # Assuming the email column is the 13th column
                cell.hyperlink = linkedin_href
                cell.value = linkedin_text
                cell.font = Font(color="0000FF", underline="single")
            workbook.save(file_path)

        print(f"URL after clicking button {index}: {current_address_url} at {current_datetime}")

        if index == 9:
            wait.until(EC.presence_of_element_located((By.XPATH, "//a[@rel='next' and @style='padding: 0.5em 0.75em' and @title='Go to next page']")))
            # Reset the index to repeat the loop when index is 10
            driver.back()
            time.sleep(5 * index)
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

