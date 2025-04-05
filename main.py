from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pandas as pd
import time

# Initialize the Firefox driver
driver = webdriver.Firefox()

# Open the website
driver.get("https://horizon.mcgill.ca/pban1/twbkwbis.P_WWWLogin")

# Wait for the user to log in
input("Did you login? Press Enter to continue...")


# Once logged in, proceed to fill the form (ONCE)
destination_city = driver.find_element(By.XPATH, '//*[@id="city_in"]')
# continue_btn = driver.find_element(By.XPATH,  '/html/body/div[3]/form/table[4]/tbody/tr[30]/td/input')

# Fill the top part of the form (ONCE)
destination_city.send_keys("Vancouver")  # Replace with your value
time.sleep(1)

province = driver.find_element(By.XPATH, '/html/body/div[3]/form/table[3]/tbody/tr[2]/td[2]/select')
province.send_keys("British Columbia")  # Replace with your value
time.sleep(2)

start_campaign_day = driver.find_element(By.XPATH, '/html/body/div[3]/form/table[3]/tbody/tr[4]/td[2]/input')
start_campaign_day.send_keys("01-Mar-2025")  # Replace with your value
time.sleep(2)


return_date = driver.find_element(By.XPATH, '/html/body/div[3]/form/table[3]/tbody/tr[5]/td[2]/input')
return_date.send_keys("18-Mar-2025")  # Replace with your value
time.sleep(2)

purpose_select = driver.find_element(By.XPATH, '/html/body/div[3]/form/table[3]/tbody/tr[6]/td[2]/select')
purpose_select.send_keys('Travel')
time.sleep(1)

describe_purpose = driver.find_element(By.XPATH, '//*[@id="te_purpose_in"]')
describe_purpose.send_keys("Measuring Something ...")  # Replace with your value
time.sleep(1)

fund_code = driver.find_element(By.XPATH, '//*[@id="fundcode_def"]')
fund_code.send_keys("44444")

time.sleep(3)
clamain_affiliation = driver.find_element(By.XPATH, '//*[@id="claim_affiliation"]')
clamain_affiliation.send_keys("Graduate Student")
time.sleep(1)
clamain_affiliation = driver.find_element(By.XPATH, '//*[@id="claim_affiliation"]')
clamain_affiliation.send_keys("Graduate Student")
time.sleep(3)
#continue_btn = driver.find_element(By.XPATH, '/html/body/div[3]/form/table[4]/tbody/tr[30]/td/input')
#continue_btn.click()
# click enter to continue
input("Continue")
time.sleep(2)

# Read the CSV file
df = pd.read_csv("20250321_SubmitExpenseReport.csv")

# Iterate over each row in the DataFrame to fill receipt-related fields
for row in df.itertuples():
    # Find the receipt-related fields for each iteration
    # //*[@id="trans_date_in"]
    receipt_date = driver.find_element(By.XPATH, '//*[@id="trans_date_in"]')

    # Fill in the receipt-related fields with data from the CSV
    receipt_date.send_keys(row.Date)
    time.sleep(1)

    expense_item = driver.find_element(By.XPATH, '//*[@id="itm_cde"]')
    expense_dropdown = Select(expense_item)
    expense_dropdown.select_by_visible_text(row.ExpenseItem) 
    time.sleep(3)

    expense_item = driver.find_element(By.XPATH, '//*[@id="itm_cde"]')
    expense_dropdown = Select(expense_item)
    expense_dropdown.select_by_visible_text(row.ExpenseItem) 
    time.sleep(2)

    description = driver.find_element(By.XPATH, '/html/body/div[3]/form/table[2]/tbody/tr[1]/td[2]/input')
    description.send_keys(row.Description)
    time.sleep(1)

    amount = driver.find_element(By.XPATH, '//*[@id="txn_amt"]')
    amount.send_keys(str(row.TotalAmount).strip()[1:])  # Ensure the amount is a string
    time.sleep(1)
    amount = driver.find_element(By.XPATH, '//*[@id="txn_amt"]')
    amount.click()

    time.sleep(2)
    purchasing_location = driver.find_element(By.XPATH, '//*[@id="taxn_code"]')
    purchasing_dropdown = Select(purchasing_location)
    purchasing_dropdown.select_by_visible_text("Canada not Quebec Location")
    time.sleep(2)
    purchasing_location = driver.find_element(By.XPATH, '//*[@id="taxn_code"]')
    purchasing_dropdown = Select(purchasing_location)
    purchasing_dropdown.select_by_visible_text("Canada not Quebec Location")


    if row.location == 'Quebec Location':
        time.sleep(2)
        purchasing_location = driver.find_element(By.XPATH, '//*[@id="taxn_code"]')
        purchasing_dropdown = Select(purchasing_location)
        purchasing_dropdown.select_by_visible_text(row.location)
        time.sleep(2)
        purchasing_location = driver.find_element(By.XPATH, '//*[@id="taxn_code"]')
        purchasing_dropdown = Select(purchasing_location)
        purchasing_dropdown.select_by_visible_text(row.location)
    time.sleep(3)
    
    account_code = driver.find_element(By.NAME, 'acctcode_in')  # Corrected XPATH for account code
    account_code.clear()
    time.sleep(2)
    account_code.send_keys("700486")
    time.sleep(1)

    add_receipt_button = driver.find_element(By.XPATH, '/html/body/div[3]/form/table[7]/tbody/tr/td[2]/input')
    add_receipt_button.click()
    time.sleep(3)

# Submit the form after all receipts are added
# Example:
# submit_button = driver.find_element(By.XPATH, '//*[@id="submit_button"]')
# submit_button.click()
# time.sleep(5)  # Wait for submission to complete

# Ensure no errors are present on the page
assert "No results found." not in driver.page_source

input("endd?")
# Close the browser
driver.close()