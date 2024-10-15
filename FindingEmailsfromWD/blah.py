import tkinter as tk
from tkinter import filedialog
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, ElementNotInteractableException
import time

def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path

def get_workday_ids(file_path):
    df = pd.read_excel(file_path)
    return df['WD ID'].tolist()

def setup_webdriver():
    options = Options()
    options.add_argument("start-maximized")  # This will start the browser maximized
    driver = webdriver.Edge(options=options)
    return driver

def navigate_to_workday(driver):
    driver.get("https://wd3.myworkday.com/sanofi/d/home.htmld")
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-automation-id='globalSearchInput']"))
    )

def wait_for_dom_stable(driver, timeout=10):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script('return document.readyState') == 'complete'
    )
    time.sleep(1)  # Additional short wait to ensure DOM is stable

def clear_search_bar(driver):
    try:
        wait_for_dom_stable(driver)
        search_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-automation-id='globalSearchInput']"))
        )
        search_input.clear()
        search_input.send_keys(Keys.ESCAPE)
        time.sleep(0.5)  # Short wait after clearing
    except Exception as e:
        print(f"Error clearing search bar: {str(e)}")

def search_workday_id(driver, workday_id):
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            clear_search_bar(driver)
            wait_for_dom_stable(driver)
            search_input = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-automation-id='globalSearchInput']"))
            )
            search_input.send_keys(workday_id)
            search_input.send_keys(Keys.RETURN)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='pexsearch-70qvj9']/a/span"))
            )
            return
        except (StaleElementReferenceException, ElementNotInteractableException):
            if attempt < max_attempts - 1:
                print(f"Element issue encountered. Retrying... (Attempt {attempt + 1})")
                time.sleep(2)
            else:
                print(f"Failed to search for {workday_id} after {max_attempts} attempts.")
                raise

def get_email_address(driver):
    try:
        wait_for_dom_stable(driver)
        email_element = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[2]/div[1]/section/div/div/div/div[1]/div/div/div/div/div/div[2]/div[2]/div[2]/ol/li/div/ol/li/div/section/div/section/div/div[4]/div/div/a/span"))
        )
        return email_element.text
    except (TimeoutException, StaleElementReferenceException):
        return "Email not found"

def main():
    file_path = select_excel_file()
    workday_ids = get_workday_ids(file_path)
    
    output_df = pd.DataFrame({'WorkdayID': workday_ids})
    output_df['Email'] = ''
    
    driver = setup_webdriver()
    
    try:
        for index, workday_id in enumerate(workday_ids):
            try:
                navigate_to_workday(driver)
                search_workday_id(driver, workday_id)
                email = get_email_address(driver)
                output_df.at[index, 'Email'] = email
                print(f"Processed {workday_id}: {email}")
                
                # Save progress after each successful processing
                output_df.to_excel('output_workday_emails_progress.xlsx', index=False)
                
                # Sleep for 10 minutes (600 seconds)
                print(f"Sleeping for 10 minutes before processing the next ID...")
                time.sleep(5)
            except Exception as e:
                print(f"Error processing {workday_id}: {str(e)}")
                output_df.at[index, 'Email'] = "Error: " + str(e)
    finally:
        driver.quit()
    
    output_file = 'output_workday_emails_final.xlsx'
    output_df.to_excel(output_file, index=False)
    print(f"Final results saved to {output_file}")

if __name__ == "__main__":
    main()