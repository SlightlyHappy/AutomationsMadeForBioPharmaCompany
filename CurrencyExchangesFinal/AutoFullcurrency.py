from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import os
import time
import datetime

# Set User-Agent for Edge options
user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
options = webdriver.EdgeOptions()
options.add_argument(f'user-agent={user_agent}')

# Get the user's home download directory
home_directory = os.path.expanduser("~")
downloads_directory = os.path.join(home_directory, "Downloads")

# Initialize the webdriver with options
driver = webdriver.Edge(options=options)

# Directly navigate to the ERIS report page
driver.get("https://eris.sanofi.com:8443/ERIS/#reportPage%26report%3DSANAVE_BMAR")

# Wait for the page to load completely
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

# Perform the tab and enter keystrokes
actions = ActionChains(driver)
actions.send_keys(Keys.TAB).send_keys(Keys.TAB).send_keys(Keys.TAB)  
actions.send_keys(Keys.ENTER)  
actions.send_keys(Keys.TAB).send_keys(Keys.ENTER) 
actions.send_keys(Keys.TAB).send_keys(Keys.ENTER) 
actions.perform()

try:
    # ... (other actions in your script)

    # 1. Locate the first button using its XPath
    first_button_xpath = "/html/body/div/div[2]/table/tbody/tr/td/table/tbody/tr/td[5]/div/div/div/div[1]/div[2]/div/div/div/div[3]/table/tbody[1]/tr/td[3]/div/div/table/tbody/tr/td/table/tbody/tr/td[1]/button"
    first_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, first_button_xpath))
    )

    # 2. Click the first button
    first_button.click()

    # ... (potential wait or other actions after the first click, if needed)

except Exception as e:
    print(f"An error occurred while clicking the first button: {e}")


try:
    # ... (continuing from the previous try block)

    # 3. Locate the second button using its XPath
    second_button_xpath = "/html/body/div/div[2]/table/tbody/tr/td/table/tbody/tr/td[5]/div/div/div/div[1]/div[2]/div/div/div/div[3]/table/tbody[1]/tr/td[3]/div/div/button"
    second_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, second_button_xpath))
    )

    # 4. Click the second button
    second_button.click()

    # ... (continue with other actions in your script)

except Exception as e:
    print(f"An error occurred while clicking the second button: {e}")


try:
    # Define the XPaths 
    download_report_button_xpath = "/html/body/div[1]/div[2]/table/tbody/tr/td/table/tbody/tr/td[5]/div/div/div/div[2]/div[1]/div/div/div[1]/table/tbody/tr/td/table/tbody/tr/td[5]/span"
    button_xpaths = [
        "/html/body/div[1]/div[2]/table/tbody/tr/td/table/tbody/tr/td[5]/div/div/div/div[1]/div[1]/div/div",
        "/html/body/div[1]/div[2]/table/tbody/tr/td/table/tbody/tr/td[5]/div/div/div/div[1]/div[2]/div/div/div/div[3]/table/tbody[1]/tr/td[3]/div/div/table/tbody/tr/td/table/tbody/tr/td[9]/button",
        "/html/body/div[1]/div[2]/table/tbody/tr/td/table/tbody/tr/td[5]/div/div/div/div[1]/div[2]/div/div/div/div[3]/table/tbody[1]/tr/td[3]/div/div/button"
    ]
    second_download_button_xpath = "/html/body/div[1]/div[2]/table/tbody/tr/td/table/tbody/tr/td[5]/div/div/div/div[2]/div[1]/div/div/div[1]/table/tbody/tr/td/table/tbody/tr/td[5]/span"

    # Create the download directories if they don't exist
    script_dir = os.path.dirname(os.path.abspath(__file__))
    year0_dir = os.path.join(script_dir, "Year0")
    year1_dir = os.path.join(script_dir, "Year1")
    os.makedirs(year0_dir, exist_ok=True)
    os.makedirs(year1_dir, exist_ok=True)

    # 1. Click the first download report button and save to Year0
    download_report_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, download_report_button_xpath))
    )
    download_report_button.click()

    # Wait for the download to complete 
    time.sleep(10)  

    # Move the latest downloaded file from the user's download directory to the Year0 folder
    downloaded_files = [f for f in os.listdir(downloads_directory) if os.path.isfile(os.path.join(downloads_directory, f))]
    if downloaded_files:
        latest_file = max(downloaded_files, key=lambda f: os.path.getctime(os.path.join(downloads_directory, f)))
        os.rename(os.path.join(downloads_directory, latest_file), os.path.join(year0_dir, latest_file))

    # 2. Click the sequence of buttons
    for button_xpath in button_xpaths:
        button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, button_xpath))
        )
        button.click()

        # Add any necessary waits or actions after each button click if needed

    # 3. Click the second download button and save to Year1
    second_download_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, second_download_button_xpath))
    )
    second_download_button.click()

    # Wait for the download to complete
    time.sleep(10)

    # Move the latest downloaded file from the user's download directory to the Year1 folder
    downloaded_files = [f for f in os.listdir(downloads_directory) if os.path.isfile(os.path.join(downloads_directory, f))]
    if downloaded_files:
        latest_file = max(downloaded_files, key=lambda f: os.path.getctime(os.path.join(downloads_directory, f)))
        os.rename(os.path.join(downloads_directory, latest_file), os.path.join(year1_dir, latest_file))

    # ... (any final actions in your script)

except Exception as e:
    print(f"An error occurred: {e}")

# Wait for 600 seconds
time.sleep(20)

# Close the browser when done
driver.quit()
