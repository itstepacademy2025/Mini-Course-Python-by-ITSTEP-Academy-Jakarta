# 1. Data harus berbentuk excel
# 2. Python berintekrasi dengan browser
# 3. Pergi ke alamat yang dituju
# 4. Handle elemen di dalam website
# 5. Automate input data secara berulang

import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Interaksi dengan file excel
wb = load_workbook(filename="Day 2/Web Automation/data.xlsx")
# print(wb) 

sheetRange = wb["Sheet1"]

# Setup webdriver
driver = webdriver.Chrome()

driver.get("https://demoqa.com/webtables")
driver.maximize_window()
driver.implicitly_wait(10)

# Mulai looping
index = 2

while index <= len(sheetRange["A"]):
    first_name = sheetRange["A" + str(index)].value
    last_name = sheetRange["B" + str(index)].value
    email = sheetRange["C" + str(index)].value
    age = sheetRange["D" + str(index)].value
    salary = sheetRange["E" + str(index)].value
    department = sheetRange["F" + str(index)].value

    # Handle add button
    add_btn = driver.find_element(By.ID, "addNewRecordButton")
    add_btn.click()

    # Check condition using try
    try:
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "registration-form-modal")))

        # Insert data by mix and match
        driver.find_element(By.ID, "firstName").send_keys(first_name)
        driver.find_element(By.ID, "lastName").send_keys(last_name)
        driver.find_element(By.ID, "userEmail").send_keys(email)
        driver.find_element(By.ID, "age").send_keys(age)
        driver.find_element(By.ID, "salary").send_keys(salary)
        driver.find_element(By.ID, "department").send_keys(department)
        driver.find_element(By.ID, "submit").click()

    except TimeoutException:
        print("Website sedang error")
        pass

    time.sleep(1)
    print(f"Data ke-{index} yaitu {first_name} terinput")
    index = index + 1

print("Data sudah terinput semua")