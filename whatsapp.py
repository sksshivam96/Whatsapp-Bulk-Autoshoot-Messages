import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Path to your WebDriver (e.g., chromedriver)
driver_path = r'C:\chromedriver-win64\chromedriver.exe'

# Load the Excel file
df = pd.read_excel('messages.xlsx')

# Add a Status column
df['Status'] = ''

# Initialize Chrome options
options = webdriver.ChromeOptions()
options.add_argument("--disable-notifications")
options.add_argument("--user-data-dir=C:\\path\\to\\chrome-profile")  # Optional: Adjust this path

# Initialize the WebDriver with ChromeOptions
driver = webdriver.Chrome(executable_path=driver_path, options=options)

driver.get('https://web.whatsapp.com')

# Wait for user to scan QR code
input('Press Enter after scanning QR code and WhatsApp Web is loaded completely')

for index, row in df.iterrows():
    phone_number = row['Mobile']
    
    message_body = row["Message"]


    try:
        # Search for the contact
        search_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
        search_box.clear()
        search_box.send_keys(phone_number)
        time.sleep(2)
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)

        # Wait until message box is clickable
        message_box = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        message_box.send_keys(message_body)
        time.sleep(10)
        message_box.send_keys(Keys.ENTER)
        time.sleep(5)

        df.at[index, 'Status'] = 'Sent'
        print(f'Message sent to {phone_number}')
    except Exception as e:
        df.at[index, 'Status'] = f'Failed: {e}'
        print(f'Failed to send message to {phone_number}: {e}')
    
    time.sleep(5)  # Delay to avoid being rate-limited by WhatsApp Web

# Close the WebDriver
driver.quit()

# Save the updated DataFrame back to the Excel file
df.to_excel('messages_with_status.xlsx', index=False)
