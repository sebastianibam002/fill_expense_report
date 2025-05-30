from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pandas as pd
import time
import numpy as np
from pyfiglet import Figlet
import os

   

f = Figlet(font='slant')
print(f.renderText('Expense Report Filler'))

filename = "input.xlsx"  # Expect this file to exist in the container at runtime
print(f'Running with {filename}')
if not os.path.exists(filename):
    raise FileNotFoundError("Expected file 'input.xlsx' not found in container.")
# Read the CSV file
df = pd.read_excel(filename, skiprows=1, sheet_name=0)

df_general = pd.read_excel(filename, sheet_name=1)

# Initialize the Firefox driver
driver = webdriver.Firefox()

# Open the website
driver.get("https://horizon.mcgill.ca/pban1/twbkwbis.P_WWWLogin")

# Wait for the user to log in
input(f"Did you login? Press enter...")


# Once logged in, proceed to fill the form (ONCE)
destination_city = driver.find_element(By.XPATH, '//*[@id="city_in"]')
# continue_btn = driver.find_element(By.XPATH,  '/html/body/div[3]/form/table[4]/tbody/tr[30]/td/input')

# Fill the top part of the form (ONCE)
destination_city.send_keys(df_general['DestinationCity'].values[0])  # Replace with your value
time.sleep(1)
province = driver.find_element(By.NAME, 'prov_state_in')
province.send_keys(df_general['Province'].values[0])  # Replace with your value
time.sleep(2)

country = driver.find_element(By.NAME, 'country_in')
country.send_keys(df_general['Country'].values[0])

time.sleep(2)

start_campaign_day = driver.find_element(By.NAME, 'start_date_in')
start_campaign_day.send_keys(df_general['Start'].dt.strftime('%d-%b-%Y').values[0])  # Replace with your value
time.sleep(2)


return_date = driver.find_element(By.NAME, 'return_date_in')
return_date.send_keys(df_general['End'].dt.strftime('%d-%b-%Y').values[0])  # Replace with your value
time.sleep(2)

purpose_select = driver.find_element(By.NAME, 'te_purpose_code_in')
purpose_select.send_keys(df_general['PurposeTrip'].values[0])
time.sleep(1)

describe_purpose = driver.find_element(By.XPATH, '//*[@id="te_purpose_in"]')
describe_purpose.send_keys(df_general['Describe'].values[0])  # Replace with your value
time.sleep(1)

fund_code = driver.find_element(By.XPATH, '//*[@id="fundcode_def"]')
fund_code.send_keys(str(df_general['FundCode'].values[0]))

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
input(f"Continue.. enter")

time.sleep(2)


# Iterate over each row in the DataFrame to fill receipt-related fields
for row in df.itertuples():
    # Find the receipt-related fields for each iteration
    # //*[@id="trans_date_in"]
    receipt_date = driver.find_element(By.XPATH, '//*[@id="trans_date_in"]')

    # Fill in the receipt-related fields with data from the CSV
    receipt_date.send_keys(row.ReceiptDate.strftime('%d-%b-%Y'))
    time.sleep(1)

    expense_item = driver.find_element(By.XPATH, '//*[@id="itm_cde"]')
    expense_dropdown = Select(expense_item)
    expense_dropdown.select_by_visible_text(row.Category) 
    time.sleep(3)

    expense_item = driver.find_element(By.XPATH, '//*[@id="itm_cde"]')
    expense_dropdown = Select(expense_item)
    expense_dropdown.select_by_visible_text(row.Category) 
    time.sleep(2)

    description = driver.find_element(By.XPATH, '/html/body/div[3]/form/table[2]/tbody/tr[1]/td[2]/input')
    description.send_keys(row.Description)
    time.sleep(1)

    amount = driver.find_element(By.XPATH, '//*[@id="txn_amt"]')
    amount.send_keys(str(row.TransactionArgentina))  # Ensure the amount is a string
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

    time.sleep(2)

    if row.Currency == 'Other':
        currency_type = driver.find_element(By.XPATH, '//*[@id="currncy_cde"]')
        currency_dropdown = Select(currency_type)
        currency_dropdown.select_by_visible_text('Other')

        time.sleep(2)
        
        currency_rate = driver.find_element(By.XPATH, '//*[@id="currncy_exch"]')
        time.sleep(2)
        currency_rate.clear()
        time.sleep(3)
        currency_rate = driver.find_element(By.XPATH, '//*[@id="currncy_exch"]')
        time.sleep(2)
        currency_rate.send_keys(str(np.round(row.ExchangeRate, 5)))
    
    time.sleep(2)

    if row.Location == 'Quebec Location':
        time.sleep(2)
        purchasing_location = driver.find_element(By.XPATH, '//*[@id="taxn_code"]')
        purchasing_dropdown = Select(purchasing_location)
        purchasing_dropdown.select_by_visible_text(row.Location)
        time.sleep(2)
        purchasing_location = driver.find_element(By.XPATH, '//*[@id="taxn_code"]')
        purchasing_dropdown = Select(purchasing_location)
        purchasing_dropdown.select_by_visible_text(row.Location)
        time.sleep(2)

    amount = driver.find_element(By.ID, 'txn_amt')
    amount.click()
    time.sleep(2)
    account_code = driver.find_element(By.NAME, 'acctcode_in')

    # Set value via JavaScript and trigger input/change event
    driver.execute_script("""
        arguments[0].value = '700486';
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, account_code)
    
    time.sleep(2)
    # /html/body/div[3]/form/table[7]/tbody/tr/td[2]/input
    add_receipt_button = driver.find_element(By.XPATH, '/html/body/div[3]/form/table[7]/tbody/tr/td[2]/input')
    driver.execute_script("arguments[0].click();", add_receipt_button)
    time.sleep(3)


# Ensure no errors are present on the page
assert "No results found." not in driver.page_source

input("Program end..")
# Close the browser
driver.close()