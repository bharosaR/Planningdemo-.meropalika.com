import re
import time
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Function to check if a string is a valid email address
def is_valid_email(email):
    try:
        # Check email format using regular expression
        email_pattern = r"^[a-z]+[\._]?[a-z 0-9]+[@]\w+[.]\w{2,3}$"
        if re.search(email_pattern, email):
            return True
        else:
            return False
    except Exception as e:
        print(e)
        return False

# Initialize Chrome WebDriver
chrome_service = ChromeService(ChromeDriverManager().install())
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

# Open the webpage
driver.get("https://planningdemo.meropalika.com/")
time.sleep(2)
try:

    # Explicitly wait for the email field to be clickable
    email_field = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, "//input[@id='Email']"))
    )

    # Explicitly wait for the password field to be present
    password_field = WebDriverWait(driver, 5).until(
    EC.visibility_of_element_located((By.XPATH, "//input[@id='Password']"))
    )

    # Set values for email and password
    email = "superadmin@gmail.com"
    password = "Softech@123"

    # Enter email and password values
    email_field.clear()
    email_field.send_keys(email)
    password_field.clear()
    password_field.send_keys(password)

    # Check if the email and password fields are empty
    if not email_field:
        print("Email field cannot be empty")
    if not password_field:
        print("Password field cannot be empty")

    # Add a sleep to see the result
    time.sleep(1)

    # Check if the email address is valid
    if is_valid_email(email):
        print("Valid email address")
    else:
        print("Invalid email address")
    time.sleep(1)

    # Click the SignIn button
    sign_in_button = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@value='लगइन गर्नुहोस्']"))
    )
    sign_in_button.click()
    time.sleep(6)

    try:
        # Scroll down the webpage
        target_y = 6000
        scroll_distance = 1000
        current_y = 0

        while current_y < target_y:
            driver.execute_script(f"window.scrollBy(0, {scroll_distance});")
            current_y += scroll_distance
            time.sleep(1)

        # Scroll up the webpage
        target_y_up = 0  # Target vertical scroll position for scrolling up
        scroll_distance = -1000  # Negative value to scroll up
        current_y = 6000  # Assuming the current vertical scroll position is at the bottom of the page

        while current_y > target_y_up:
                driver.execute_script(f"window.scrollBy(0, {scroll_distance});")
                current_y += scroll_distance  # Decrementing the current_y position
                time.sleep(1)

    except Exception as e:
        print(f"An error occurred: {e}")

    # Navigate to LBIS_excel
    LBIS_excel = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='LMBIS Excel']"))
    )
    LBIS_excel.click()
    time.sleep(1)

    # Locate the dropdown option and wait for it to be clickable
    # Navigate to LBIS_excel_upload
    LBIS_excel_upload = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='LMBIS Excel Upload']"))
    )
    LBIS_excel_upload.click()
    time.sleep(1)

    #Navigate to ब.उ.शि.नं.
    dropdown_element1 = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable((By.XPATH, "//b[@role='presentation']"))
    )
    dropdown_element1.click()
    time.sleep(1)

    #Select dropdown element from option
    select_element1 = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable((By.XPATH, "//option[contains(text(),'भवन निर्माण संहिता, सरकारी भवन निर्माण[34701105]')]"))
    )
    select_element1.click()
    time.sleep(1)

    #Navigate to अफिसको नाम
    dropdown_element2 = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable((By.XPATH, "//span[@id='select2-OfficeId-container']"))
    )
    dropdown_element2.click()
    time.sleep(1)

    # Select dropdown element from option
    select_element2 = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//option[contains(text(),'सघन शहरी तथा भवन निर्माण आयोजना झापा[३४७०११२०१]')]"))
    )
    select_element2.click()
    time.sleep(2)


    def english_to_nepali_number(english_number):
        english_to_nepali = {
            '0': '०', '1': '१', '2': '२', '3': '३', '4': '४',
            '5': '५', '6': '६', '7': '७', '8': '८', '9': '९'
        }
        nepali_number = ''.join(
            english_to_nepali[digit] if digit in english_to_nepali else digit for digit in english_number)
        return nepali_number


    # Get today's date in YYYY-MM-DD format
    today = datetime.datetime.today().strftime('%Y-%m-%d')

    # Convert today's date to Nepali format
    nepali_date = english_to_nepali_number(today)

    # Wait for the calendar input field to be present
    calendar_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//input[@id="YojanaStartDate"]'))
    )

    # Clear the field if needed and input the Nepali date
    calendar_field.clear()
    calendar_field.send_keys(nepali_date)
    time.sleep(2)


    #Add todays date in calender
    # Locate the calendar input field
    #calendar_field = driver.find_element(By.XPATH,
    #'//input[@id= "YojanaStartDate"]') # Replace 'calendar' with the actual ID or use other locators

    # Get today's date
    #today = datetime.datetime.today().strftime('%Y-%m-%d')  # Adjust format as needed

    # Clear the field if needed and input today's date
    #calendar_field.clear()
    #calendar_field.send_keys(today)
    #time.sleep(2)


    #Navigate to योजना सुरु मिति:
    #Date_plan = WebDriverWait(driver, 2).until(
    #   EC.element_to_be_clickable((By.XPATH, "//input[@id='YojanaStartDate']"))
    #)
    #Date_plan.click()
    #time.sleep(1)

    #Locate year
    #year_field = WebDriverWait(driver, 2).until(
     #   EC.element_to_be_clickable((By.XPATH,"//option[@value='2078']"))
    #)
    #year_field.click()
    #time.sleep(1)

    #Locate month
    month_field = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable((By.XPATH, "//option[contains(text(),'जेठ')]"))
    )
    month_field.click()
    time.sleep(1)

    # Locate day
    day_field = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable((By.XPATH, "//body[1]/div[3]/div[2]/table[1]/tbody[1]/tr[6]/td[7]/a[1]"))
    )
    day_field.click()
    time.sleep(1)

    #Navigate to योजना अन्त्य मिति
    Date_plan_end = WebDriverWait(driver, 2).until(
       EC.element_to_be_clickable((By.XPATH, "//input[@id='YojanaEndDate']"))
    )
    Date_plan_end.click()
    time.sleep(1)

    #Locate year_end
    year_field_end = WebDriverWait(driver, 2).until(
       EC.element_to_be_clickable((By.XPATH, "//option[@value='2078']"))
    )
    year_field_end.click()
    time.sleep(1)

    #Locate month_end
    month_field_end = WebDriverWait(driver, 2).until(
       EC.element_to_be_clickable((By.XPATH, "//option[contains(text(),'अषाढ')]"))
    )
    month_field_end.click()
    time.sleep(1)

    # Locate day_end
    day_field_end = WebDriverWait(driver, 2).until(
       EC.element_to_be_clickable((By.XPATH, "//body[1]/div[3]/div[2]/table[1]/tbody[1]/tr[6]/td[3]/a[1]"))
    )
    day_field_end.click()
    time.sleep(1)

    #Navigate to फाइल छान्नुहोस
    # Set path to my Excel file
    file_path = "C:/Users/ASUS/OneDrive/Desktop/MAnual test for Likhupe.xlsx"

    choose_file = WebDriverWait(driver, 2).until(
       EC.presence_of_element_located((By.XPATH, "//input[@name='file']"))
    )
    choose_file.send_keys(file_path)
    time.sleep(5)

    #Navigate to Submit Button
    Submit_button = WebDriverWait(driver,2).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@id='Submit']"))
    )
    Submit_button.click()
    time.sleep(1)

except TimeoutException as e:
    print(f"TimeoutException: {e}")
except NoSuchElementException as e:
    print(f"NoSuchElementException: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")

finally:
    # Close the WebDriver
    driver.quit()
    print("Congrats!! Code ran successfully")

