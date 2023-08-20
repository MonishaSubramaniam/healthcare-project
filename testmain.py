from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import locators
from excel_fun import Excel_Functions
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

file = 'C:\\project file\\pom1\\test_data.xlsx'
sheet_number = 'Sheet1'
s = Excel_Functions(file, sheet_number)
driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()
driver.get("https://katalon-demo-cura.herokuapp.com/")
driver.find_element(By.ID,locators.Locators().make_apnt).click()
driver.implicitly_wait(10)
rows = s.row_count()

# Read the Excel file using Rows and Columns
for row in range(2,rows+1):
    username = s.read_data(row, 6)
    password = s.read_data(row, 7)
    
# Selenium Automation Code
try:
   wait=WebDriverWait(driver,15)
   wait.until(EC.presence_of_element_located((By.ID,locators.Locators().username_locator))).send_keys(username)
   wait.until(EC.presence_of_element_located((By.ID,locators.Locators().password_locator))).send_keys(password)
   wait.until(EC.presence_of_element_located((By.ID,locators.Locators().submit_button))).click()
   wait.until(EC.presence_of_element_located((By.XPATH,locators.Locators().date))).send_keys("19/08/2023")
   wait.until(EC.presence_of_element_located((By.ID,locators.Locators().com))).send_keys("Regular Checkup")
   wait.until(EC.element_to_be_clickable((By.ID,locators.Locators().Book))).click()
except TimeoutException:
    print('element missing')

# Check for the TEST PASS or TEST FAIL

if 'https://katalon-demo-cura.herokuapp.com/#appointment' in driver.current_url:
    print('SUCCESS : login with {a}'.format(a=username))
    s.write_data(row, 8, "TEST PASS")
    driver.back()
elif 'https://katalon-demo-cura.herokuapp.com/profile.php#login' in driver.current_url:
    print("FAIL : login failure with {a}".format(a=username))
    s.write_data(row, 8, "TEST FAIL")
    driver.back()


# Quit/Close the WebDriver
driver.quit()