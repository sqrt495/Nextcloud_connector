from selenium import webdriver
from selenium.webdriver.common.by import By
import os
from time import sleep
import config

# TODO: add funcs description
# TODO: rewrite driver
def open_url_in_chrome(url, password=''):
    '''main func'''
    driver = webdriver.Chrome(config.driver_source)
    driver.implicitly_wait(15)
    driver.get(url)

    if password != '':
        driver.find_element(By.ID, 'password').send_keys(password)
        driver.find_element(By.ID, 'password-submit').click()
    return driver

def open_url_in_chrom_for_loading(url, path_for_download, password=''):
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': path_for_download}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(config.driver_source, chrome_options=chrome_options)
    driver.implicitly_wait(15)
    driver.get(url)

    if password != '':
        driver.find_element(By.ID, 'password').send_keys(password)
        driver.find_element(By.ID, 'password-submit').click()
    return driver


def download(url, path_for_download, password=''):
    driver = open_url_in_chrom_for_loading(url, path_for_download, password)
    driver.find_element(By.XPATH, '//*[@id="headerSelection"]/label').click()
    driver.find_element(By.CLASS_NAME, 'actions-selected').click()
    driver.implicitly_wait(15)
    driver.find_element(By.XPATH, '//*[@id="selectedActionsList"]/div/ul/li[2]/a').click()
    sleep(10)

    while any([filename.endswith(".crdownload") for filename in os.listdir(path_for_download)]):
        sleep(1)
    driver.quit()

def clear_cloud(url, password=''):
    driver = open_url_in_chrome(url, password)
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

def upload(url, files, password=''):
    driver = open_url_in_chrome(url, password)
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

def create_folder(url, folders_name_list, password=''):
    driver = open_url_in_chrome(url, password)
    for n in folders_name_list:
        try:
            driver.find_element(By.CLASS_NAME, 'new').click()
            driver.find_element(By.XPATH, '//*[@id="controls"]/div[2]/div[2]/ul/li[2]/a').click()
            driver.find_element(By.XPATH, '//*[@id="view5-input-folder"]').send_keys(n)
            driver.find_element(By.XPATH, '//*[@id="controls"]/div[2]/div[2]/ul/li[2]/a/form/input[2]').click()
            sleep(3)
            print(f'{n} folder created')
        except Exception as e:
            print('settings environments dirs error:', e, sep='\n')
    driver.quit()

def create_folder_and_sub_folder(url, folder_name, subfolder_name, password=''):
    """description"""
    driver = open_url_in_chrome(url, password)
    try:
        driver.find_element(By.CLASS_NAME, 'new').click()
        driver.find_element(By.XPATH, '//*[@id="controls"]/div[2]/div[2]/ul/li[2]/a').click()
        driver.find_element(By.XPATH, '//*[@id="view5-input-folder"]').send_keys(folder_name)
        driver.find_element(By.XPATH, '//*[@id="controls"]/div[2]/div[2]/ul/li[2]/a/form/input[2]').click()
        sleep(2)
        # check this

        driver.get(url+'%2F'+folder_name)
        driver.find_element(By.CLASS_NAME, 'new').click()
        driver.find_element(By.XPATH, '//*[@id="controls"]/div[2]/div[2]/ul/li[2]/a').click()
        driver.find_element(By.XPATH, '//*[@id="view5-input-folder"]').send_keys(subfolder_name)
        driver.find_element(By.XPATH, '//*[@id="controls"]/div[2]/div[2]/ul/li[2]/a/form/input[2]').click()
        sleep(3)
    except Exception as e:
        print('settings environments dirs error:', e, sep='\n')
    driver.quit()

def download_clear_create_folder(url, path_for_download, folders_name_list, password=''):
    """description"""
    # download seattings
    driver = open_url_in_chrom_for_loading(url, path_for_download, password)
    # download
    driver.find_element(By.XPATH, '//*[@id="headerSelection"]/label').click()
    driver.find_element(By.CLASS_NAME, 'actions-selected').click()
    driver.implicitly_wait(15)
    driver.find_element(By.XPATH, '//*[@id="selectedActionsList"]/div/ul/li[2]/a').click()
    sleep(10)

    while any([filename.endswith(".crdownload") for filename in os.listdir(path_for_download)]):
        sleep(3)
    print('download competed')

    # clear
    driver.refresh()
    driver.find_element(By.XPATH, '//*[@id="headerSelection"]/label').click()
    driver.find_element(By.CLASS_NAME, 'actions-selected').click()
    driver.find_element(By.XPATH, '//*[@id="selectedActionsList"]/div/ul/li[3]/a').click()
    sleep(5)

    up_bar = ''
    while up_bar != 'none':
        up_bar = driver.find_element(By.ID, 'uploadprogressbar').value_of_css_property('display')
        sleep(1)
    sleep(2)
    print('clear competed')

    # create
    for n in folders_name_list:
        try:
            driver.find_element(By.CLASS_NAME, 'new').click()
            driver.find_element(By.XPATH, '//*[@id="controls"]/div[2]/div[2]/ul/li[2]/a').click()
            driver.find_element(By.XPATH, '//*[@id="view5-input-folder"]').send_keys(n)
            driver.find_element(By.XPATH, '//*[@id="controls"]/div[2]/div[2]/ul/li[2]/a/form/input[2]').click()
            sleep(3)
            print('successfully created:',n, sep='\n')
        except Exception as e:
            print('error while creating:', n, sep='\n')
    print('create competed')
    driver.quit()

