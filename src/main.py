import time
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains



# service = Service(driver_location)
driver = webdriver.Firefox()

driver.get('https://ti2ptih6lthnanlb.vercel.app/')


# sudo mv chromedriver /usr/local/share/chromedriver
dataframe = pd.read_excel('data.xlsx')
# Loop through each row in the dataframe
for index, entry in dataframe.iterrows():
    print(f"Processing entry {index + 1} of {len(dataframe)}")
    
    # Wait for form to be ready
    wait = WebDriverWait(driver, 10)
    name_input = wait.until(EC.presence_of_element_located((By.ID, 'fullName')))
    
    # Fill form fields
    name_input.send_keys(entry['Full Name'])

    passport_input = driver.find_element(By.ID,'passportNumber')
    passport_input.send_keys(str(entry['PassportNumber']))

    nationality_input = driver.find_element(By.ID,'nationality')
    nationality_input.send_keys(entry['Nationality'])

    email_input = driver.find_element(By.ID,'email')
    email_input.send_keys(entry['Email'])

    # Set visa type
    x = entry['VisaType'].lower()
    driver.execute_script(f"""
        const event = new Event('change', {{ bubbles: true }});
        const select = document.querySelector('select');
        select.value = '{x}';
        select.dispatchEvent(event);
    """)

    # Wait a bit for any animations/state updates
    time.sleep(1)

    # Find and click the submit button
    submit_button = driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]')
    driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
    time.sleep(0.5)  # Wait for scroll
    submit_button.click()

    # Wait for the success message
    wait = WebDriverWait(driver, 10)
    success_message = wait.until(
        EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Thank You for Your Submission')]"))
    )

    # Wait a moment to see the success message
    time.sleep(2)

    # Navigate back to the form page if not the last entry
    if index < len(dataframe) - 1:  # Only go back if there are more entries to process
        driver.back()
        
        # Wait for the form to load again
        wait.until(EC.presence_of_element_located((By.ID, 'fullName')))
        time.sleep(1)  # Additional wait to ensure form is fully loaded

print("All entries have been processed!")
# Optional: Close the browser
# driver.quit()