from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import csv



def generate_data(driver):
    stories = driver.find_element(By.CLASS_NAME, 'grid-table-wrapper').find_element(By.CLASS_NAME, 'main').find_elements(By.TAG_NAME, 'tr')
    for story in stories:
        sprint_list = list()
        for item in story.find_elements(By.CLASS_NAME, 'subtable'):
            try:
                sprint_list.append(item.text)
            except:
                sprint_list.append(item.find_element(By.TAG_NAME, 'a').text)
        yield sprint_list


def packageData(driver):
    with open('data.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        for story in generate_data(driver):
            writer.writerow(story)
    
    


def start():
    driver = webdriver.Chrome(ChromeDriverManager().install())
    #driver.get("https://www6.v1host.com/ToshibaGCS/TeamRoom.mvc/Show/787606")
    #driver.get("https://www6.v1host.com/ToshibaGCS/Schedule.mvc/Summary?oidToken=Schedule%3A782414")
    driver.get('https://www6.v1host.com/ToshibaGCS/Default.aspx?menu=PrimaryBacklogPage')

    #wait = WebDriverWait(driver, 10)

    driver.implicitly_wait(10)
    username_field = driver.find_element(By.ID, 'username')
    password_field = driver.find_element(By.ID, 'password')

    driver.implicitly_wait(10)
    submit_button = driver.find_element(By.CLASS_NAME, 'confirm-btn')

    username_field.send_keys("seth.wiggins")
    password_field.send_keys("Welcome#01")
    submit_button.click()

    driver.implicitly_wait(10)
    time.sleep(2)
    packageData(driver)
    driver.close()
