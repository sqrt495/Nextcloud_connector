from selenium import webdriver
from selenium.webdriver.common.by import By
import os
from time import sleep
import telegram
import config

def download(url, path, password=''):
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': path}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(config.driver_source, chrome_options=chrome_options)
    driver.implicitly_wait(15)
    driver.get(url)

    if password != '':
        driver.find_element(By.ID, 'password').send_keys(password)
        driver.find_element(By.ID, 'password-submit').click()

    driver.find_element(By.XPATH, '//*[@id="headerSelection"]/label').click()
    driver.find_element(By.CLASS_NAME, 'actions-selected').click()
    driver.implicitly_wait(15)
    driver.find_element(By.XPATH, '//*[@id="selectedActionsList"]/div/ul/li[2]/a').click()

    sleep(10)

    while any([filename.endswith(".crdownload") for filename in os.listdir(path)]):
        sleep(1)

    driver.quit()


def upload(url, files, password=''):
    driver = webdriver.Chrome(config.driver_source)
    driver.implicitly_wait(15)
    driver.get(url)

    if password != '':
        driver.find_element(By.ID, 'password').send_keys(password)
        driver.find_element(By.ID, 'password-submit').click()

    driver.implicitly_wait(15)

    for file in files:
        driver.find_element(By.XPATH, '//*[@id="file_upload_start"]').send_keys(file)

    sleep(5)

    up_bar = ''

    while up_bar != 'none':
        up_bar = driver.find_element(By.ID, 'uploadprogressbar').value_of_css_property('display')
        sleep(1)

    sleep(2)
    driver.quit()


def create_folder(url, name, password=''):
    driver = webdriver.Chrome(config.driver_source)
    driver.implicitly_wait(15)
    driver.get(url)

    if password != '':
        driver.find_element(By.ID, 'password').send_keys(password)
        driver.find_element(By.ID, 'password-submit').click()

    driver.find_element(By.CLASS_NAME, 'new').click()
    driver.find_element(By.XPATH, '//*[@id="controls"]/div[2]/div[2]/ul/li[2]/a').click()
    driver.find_element(By.XPATH, '//*[@id="view5-input-folder"]').send_keys(name)
    driver.find_element(By.XPATH, '//*[@id="controls"]/div[2]/div[2]/ul/li[2]/a/form/input[2]').click()

    sleep(3)
    driver.quit()


def clear_cloud(url, password=''):
    driver = webdriver.Chrome(config.driver_source)
    driver.implicitly_wait(15)
    driver.get(url)

    if password != '':
        driver.find_element(By.ID, 'password').send_keys(password)
        driver.find_element(By.ID, 'password-submit').click()

    driver.find_element(By.XPATH, '//*[@id="headerSelection"]/label').click()
    driver.find_element(By.CLASS_NAME, 'actions-selected').click()
    driver.find_element(By.XPATH, '//*[@id="selectedActionsList"]/div/ul/li[3]/a').click()

    sleep(5)

    up_bar = ''

    while up_bar != 'none':
        up_bar = driver.find_element(By.ID, 'uploadprogressbar').value_of_css_property('display')
        sleep(1)

    sleep(2)
    driver.quit()


def send_message_bot(msg):
    bot = telegram.Bot(token=config.bot_token, request=telegram.utils.request.Request())
    bot.sendMessage(chat_id=config.chat_id, text=msg)
    bot.sendMessage(chat_id=config.chat_id2, text=msg)
    bot.sendMessage(chat_id=config.chat_id3, text=msg)
