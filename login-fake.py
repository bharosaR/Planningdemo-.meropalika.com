from selenium import webdriver
from selenium.webdriver.common.by import By
from faker import Faker

import time

# Create a Faker instance
fake = Faker()

# Initialize WebDriver
driver = webdriver.Chrome()
driver.get("https://planningdemo.meropalika.com/")

email_input = driver.find_element(By.XPATH, "//input[@id='Email']")
email_input.send_keys(fake.email())

password_input = driver.find_element(By.XPATH, "//input[@id='Password']")
password_input.send_keys(fake.password())
time.sleep(1)

# Submit the form
submit_button = driver.find_element(By.XPATH, "//input[@value='लगइन गर्नुहोस्']")
submit_button.click()
time.sleep(5)

# Close the driver
driver.quit()

