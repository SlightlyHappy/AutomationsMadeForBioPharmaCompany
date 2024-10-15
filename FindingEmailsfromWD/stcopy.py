import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import logging

# Setup logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class AutomationApp:
    def __init__(self, master):
        self.master = master
        self.master.title('Workday Automation')
        self.master.geometry("500x300")
        self.master.config(background="white")

        self.excel_file = ""
        self.setup_ui()

    def setup_ui(self):
        self.label_file = tk.Label(self.master, text="No file selected", width=50, height=4, fg="blue")
        self.label_file.pack(pady=10)

        self.button_browse = tk.Button(self.master, text="Browse Files", command=self.browse_files)
        self.button_browse.pack(pady=10)

        self.button_run = tk.Button(self.master, text="Run Automation", command=self.run_automation)
        self.button_run.pack(pady=10)

        self.button_exit = tk.Button(self.master, text="Exit", command=self.master.quit)
        self.button_exit.pack(pady=10)

    def browse_files(self):
        # File selection dialog
        self.excel_file = filedialog.askopenfilename(
            initialdir="/",
            title="Select an Excel File",
            filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*"))
        )
        if self.excel_file:
            self.label_file.config(text=f"File Selected: {self.excel_file}")

    def wait_for_element(self, driver, by, value, timeout=20):
        try:
            return WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
        except TimeoutException:
            logging.error(f"Timeout waiting for element {value}")
            raise

    def run_automation(self):
        if not self.excel_file:
            messagebox.showerror("Error", "Please select an Excel file first.")
            return

        try:
            df = pd.read_excel(self.excel_file)
            driver = webdriver.Edge()
            driver.get("https://wd3.myworkday.com/sanofi/d/home.htmld")

            # Open the Workday login page
            driver.get("https://wd3.myworkday.com/sanofi/d/home.htmld")

            # Wait for the login to be completed (if needed) - Adjust logic as necessary
            time.sleep(10)  # Add login logic here if needed

            # Loop through Workday IDs and search for emails
            for index, row in df.iterrows():
                wd_id = row['WD ID']
                logging.info(f"Processing WD ID: {wd_id}")
                
                try:
                    # Wait for and locate the global search input field
                    search_input = self.wait_for_element(
                        driver,
                        By.CSS_SELECTOR,
                        "input[data-automation-id='globalSearchInput']"
                    )
                    
                    # Clear search bar, input Workday ID, and press enter
                    search_input.clear()
                    search_input.send_keys(wd_id)
                    time.sleep(1)
                    search_input.send_keys(Keys.RETURN)
                    logging.info(f"Search initiated for WD ID: {wd_id}")

                    # Wait for page to load results
                    time.sleep(5)

                    # Capture email address using the provided XPath
                    try:
                        email_element = self.wait_for_element(
                            driver,
                            By.XPATH,
                            "/html/body/div[2]/div/div[2]/div[1]/section/div/div/div/div[1]/div/div/div/div/div/div[2]/div[2]/div[2]/ol/li/div/ol/li/div/section/div/section/div/div[4]/div/div/a/span"
                        )
                        email_address = email_element.get_attribute("textContent").strip()
                        logging.info(f"Email found for WD ID {wd_id}: {email_address}")
                    except (TimeoutException, NoSuchElementException):
                        email_address = "NA"
                        logging.warning(f"Email not found for WD ID: {wd_id}")

                except TimeoutException:
                    email_address = "NA"
                    logging.error(f"Timeout occurred while processing WD ID: {wd_id}")

                except Exception as e:
                    email_address = "NA"
                    logging.error(f"Error processing WD ID {wd_id}: {str(e)}")

                # Save email to the dataframe
                df.loc[index, 'Email address'] = email_address

                # Wait a few seconds before processing the next ID
                time.sleep(5)

            # Save results to a new Excel file
            output_file = 'output_emails.xlsx'
            df.to_excel(output_file, index=False)
            logging.info(f"Automation completed. Results saved to {output_file}")
            driver.quit()

            # Notify the user that the process is completed
            messagebox.showinfo("Success", f"Automation completed. Results saved to {output_file}")

        except Exception as e:
            logging.error(f"An error occurred during automation: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomationApp(root)
    root.mainloop()
