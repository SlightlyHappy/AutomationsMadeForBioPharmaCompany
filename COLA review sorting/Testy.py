import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import os
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains


def login(driver, username, password):
    try:
        # Wait for the initial username field and enter the username
        initial_username_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "useremailaddress"))
        )
        initial_username_field.send_keys(username)
        initial_username_field.send_keys(Keys.RETURN)  # Submit

        # Wait for the second username field and check if it's empty
        username_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "txtEmail"))
        )
        if not username_field.get_attribute("value"):  # Check if empty
            username_field.send_keys(username)
            username_field.send_keys(Keys.RETURN)  # Submit

        # Wait for the password field, click it, and enter the password
        password_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "txtPassword"))
        )
        password_field.click()
        password_field.send_keys(password)
        password_field.send_keys(Keys.RETURN)  # Submit

    except NoSuchElementException:
        print(f"Element not found. Check your selectors.")
    except Exception as e:
        print(f"An error occurred: {e}")


def click_accept_button(driver):
    try:
        # Wait for the accept button to be clickable and then click it
        accept_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "truste-consent-button"))
        )
        accept_button.click()
    except NoSuchElementException:
        print("Accept button not found.")
    except Exception as e:
        print(f"An error occurred while clicking the accept button: {e}")


def click_cost_of_living_link(driver):
    try:
        # Click the "Accept" button if it exists
        click_accept_button(driver)

        # Explicit wait for link clickability
        link_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Cost of Living Calculator"))  # replace with XPATH for dynamic clicking
        )
        link_element.click()  # Click the link

        # Implicit wait for new window to open
        driver.implicitly_wait(10)

        # Print current window handles for debugging
        print(driver.window_handles)

        # Switch to the new window
        all_windows = driver.window_handles
        driver.switch_to.window(all_windows[-1])

        # Optionally, add a wait to ensure the new page has loaded completely
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        # Handle the additional consent button
        consent_button_xpath = "/html/body/div[1]/div/div[3]/div[3]/div/button[2]"
        consent_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, consent_button_xpath))
        )
        consent_button.click()

        # Optionally, add a wait to ensure the new page has loaded completely
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    except NoSuchElementException:
        print("Cost of Living Calculator 2.0 link or consent button not found.")
    except Exception as e:
        print(f"An error occurred while clicking the link or consent button: {e}")

# Set up the Edge webdriver
service = Service(EdgeChromiumDriverManager().install())
driver = webdriver.Edge(service=service)

# Replace with the actual website you want to visit
website_url = "https://mobilityexchange.mercer.com/dashboard/MyTools"
driver.get(website_url)

# Read credentials from a text file
script_dir = os.path.dirname(os.path.abspath(__file__))
txt_file_path = os.path.join(script_dir, "Credentials.txt")

with open(txt_file_path, 'r') as file:
    lines = file.readlines()

# Parse username and password from file
username = lines[0].split('=')[1].strip().strip("'")
password = lines[1].split('=')[1].strip().strip("'")

def browse_file():
    global excel_file_path  # Make excel_file_path accessible outside the function
    file_path = filedialog.askopenfilename(
        initialdir=os.path.dirname(os.path.abspath(__file__)),
        title="Select Excel File",
        filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")),
    )
    if file_path:
        excel_file_path = file_path
        file_label.config(text=file_path)



def select_option_from_excel():
    global excel_file_path, df  # Access global variables

    if excel_file_path:
        try:
            excel_df = pd.read_excel(excel_file_path)
            print("Excel file read successfully.")

            if 'Home Country, City' in excel_df.columns and 'Host Country, City' in excel_df.columns:
                home_cities = excel_df['Home Country, City']
                host_cities = excel_df['Host Country, City']
                print("Columns 'Home Country, City' and 'Host Country, City' found.")
            else:
                print("Error: One or both required columns not found in the Excel file.")
                return

            # Force string conversion for both columns
            home_cities = home_cities.astype(str).str.strip().str.lower()
            host_cities = host_cities.astype(str).str.strip().str.lower()
            df['Text'] = df['Text'].astype(str).str.strip().str.lower()

            # Iterate through Excel rows and select options
            for i in range(len(home_cities)):
                home_city = home_cities.iloc[i]
                host_city = host_cities.iloc[i]

                # Select Home City
                matching_row_home = df[df['Text'] == home_city]
                if not matching_row_home.empty:
                    row_index_home = matching_row_home.index[0]
                    select_home_element = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.ID, 'locations-home-location'))
                    )
                    select_home_element.click()
                    for _ in range(row_index_home):
                        actions = ActionChains(driver)
                        actions.send_keys(Keys.ARROW_DOWN).perform()
                    actions.send_keys(Keys.ENTER).perform()
                    print(f"Selected Home City option for '{home_city}' at row index {row_index_home}")
                else:
                    print(f"No match found for Home City '{home_city}' in the DataFrame.")

                # Select Host City (adjusting the index and adding potential fixes)
                matching_row_host = df[df['Text'] == host_city]
                if not matching_row_host.empty:
                    row_index_host = matching_row_host.index[0] - 2

                    # Re-locate the dropdown element
                    select_host_element = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, '/html/body/mercer-app/div/main/mmp-cost-of-living-calculator-container/mmp-col-start-container/mmp-col-start/form/mercer-stepper/div/div[4]/mercer-step[1]/mmp-locations/mercer-card/div/div[3]/mercer-card-content[2]/div[2]/div[2]/div/mmp-select/div/div[1]/div/select'))
                    )

                    # Add a small delay to ensure the dropdown is fully open
                    select_host_element.click()
                    time.sleep(1)

                    # Try selecting by visible text first
                    try:
                        select = Select(select_host_element)
                        select.select_by_visible_text(host_city) 
                        print(f"Selected Host City option for '{host_city}' using select_by_visible_text")
                    except NoSuchElementException:
                        # If select_by_visible_text fails, fallback to key presses
                        for _ in range(row_index_host):
                            actions = ActionChains(driver)
                            actions.send_keys(Keys.ARROW_DOWN).perform()
                        actions.send_keys(Keys.ENTER).perform()
                        print(f"Selected Host City option for '{host_city}' at row index {row_index_host + 1}")
                else:
                    print(f"No match found for Host City '{host_city}' in the DataFrame.")

                # Click the button 
                button_element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'RUN A NEW CALCULATION')]"))
                )
                button_element.click()

                # Wait for the next page to load (adjust the wait as needed)
                time.sleep(5)

                # Perform actions on the next page using data from the same Excel row
                # ... (Add your code here to interact with elements on the next page)

        except FileNotFoundError:
            print(f"Error: File '{excel_file_path}' not found.")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
    else:
        print("Please select an Excel file first.")


try:
    # Login
    login(driver, username, password)

    # Click on the link after login
    click_cost_of_living_link(driver)

    # Wait for the calculator page to load
    time.sleep(5)

    # Locate the 'select' element and get its outerHTML
    select_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'locations-home-location'))
    )
    html_content = select_element.get_attribute('outerHTML')

    # Parse the HTML and extract data
    soup = BeautifulSoup(html_content, 'html.parser')

    # Extract data and create DataFrame
    data = []
    for option in soup.find_all('option'):
        if option.get('value') and not option.get('disabled'):
            data.append({
                'Option Value': option['value'],
                'Text': option.text.strip()
            })

    df = pd.DataFrame(data)

    # Print the entire DataFrame
    print(df)

    # Create the main window for file selection (after printing the DataFrame)
    root = tk.Tk()
    root.title("Select Excel File")

    # Browse button
    browse_button = tk.Button(root, text="Browse Excel File", command=browse_file)
    browse_button.pack(pady=10)

    # File label
    file_label = tk.Label(root, text="No file selected")
    file_label.pack()

    # Select option button
    select_button = tk.Button(root, text="Select Option from Excel", command=select_option_from_excel)
    select_button.pack(pady=10)

    # Run the GUI
    root.mainloop()

    # Optionally, add a delay to observe the result before closing the browser
    time.sleep(500)

except Exception as e:
    print(f"An error occurred: {e}")
finally:
    # Close the browser
    driver.quit()
