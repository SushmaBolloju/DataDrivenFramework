from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from locators import Test_Locators
from excelfunctions import Excel_Functions
from webdriver_manager.chrome import ChromeDriverManager

excel_file = "/data2.csv"

sheet_number = "Sheet1"

s = Excel_Functions(excel_file, sheet_number)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
# driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))

driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")

rows = Excel_Functions(excel_file, sheet_number).Row_Count()

for row in range(2, rows+1):
    username = Excel_Functions(excel_file, sheet_number).Read_Data(row, 6)
    password = Excel_Functions(excel_file, sheet_number).Read_Data(row, 7)

    driver.find_element(by=By.XPATH, value=Test_Locators().username_locator).send_keys(username)
    driver.find_element(by=By.XPATH, value=Test_Locators().password_locator).send_keys(password)
    driver.find_element(by=By.XPATH, value=Test_Locators().submitButton_locator).click()

     # Wait for the page to load and check if login was successful
    driver.implicitly_wait(10)
    if 'https://opensource-demo.orangehrmlive.com/web/index.php/auth/login/checkpoint/?next' in driver.current_url:
        print("SUCCESS : Login Success with Username {a}".format(a = username))
        Excel_Functions(excel_file, sheet_number).Write_Data(row,8, "TEST PASS")
        driver.back()
    elif('https://opensource-demo.orangehrmlive.com/web/index.php/auth/login' in driver.current_url):
        print("FAIL : Login Failure with Username {a}".format(a = username))
        Excel_Functions(excel_file, sheet_number).Write_Data(row,8, "TEST FAIL")
        driver.back()

# Close the browser
driver.quit()