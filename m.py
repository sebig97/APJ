import os
import glob
import time
import pandas
import datetime
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options


class homePage(object):
    PATH_TO_CHROME_DRIVER = r"C:\chromedriver\chromedriver.exe"

    def __init__(self, base_url, merchant, username, password):
        self.base_url = base_url
        self.merchant = merchant
        self.username = username
        self.password = password
        # self.driver = webdriver

    def get_chrome_driver(self):
        chrome_options = Options()
        prefs = {
            "download.default_directory": r"C:\DownloadFiles",
            "safebrowsing.enabled": "false"}
        chrome_options.add_experimental_option("prefs", prefs)

        self.driver = webdriver.Chrome(executable_path=homePage.PATH_TO_CHROME_DRIVER, options=chrome_options)

        # browser should be loaded in maximized window
        self.driver.maximize_window()

        # driver should wait implicitly for a given duration, for the element under consideration to load.
        self.driver.implicitly_wait(10)  # 10 is in seconds

        print("chrome browser successfully opened")

    # --- Steps ---
    def test_load_home_page(self):
        # to initialize a variable to hold reference of webdriver instance being passed to the function as a reference.
        driver = self.driver
        # to load a given URL in browser window
        driver.get(self.base_url)
        # time.sleep(3)

        print("page successfully loaded")

    def login_veritrans(self):
        driver = self.driver
        print(driver.current_url)

        # login credentials
        driver.find_element_by_xpath("//a[contains(text(),'For User')]").click()
        search_merchant = driver.find_element_by_xpath("//div[3]//form[1]//table[1]//tbody[1]//tr[1]//td[1]//input[1]")
        search_merchant.clear()
        search_merchant.send_keys(self.merchant)

        search_user_id = driver.find_element_by_xpath("//div[3]//form[1]//table[1]//tbody[1]//tr[2]//td[1]//input[1]")
        search_user_id.clear()
        search_user_id.send_keys(self.username)

        search_password = driver.find_element_by_xpath("//div[3]//form[1]//table[1]//tbody[1]//tr[3]//td[1]//input[1]")
        search_password.clear()
        search_password.send_keys(self.password)
        search_password.send_keys(Keys.RETURN)

        print("Logging successfully to VERITRANS page")
        # time.sleep(2)

    def login_JCB(self):
        driver = self.driver
        print(driver.current_url)

        # login credentials
        search_merchant = driver.find_element_by_id("loginId")
        search_merchant.clear()
        search_merchant.send_keys(self.username)

        search_password = driver.find_element_by_id("password")
        search_password.clear()
        search_password.send_keys(self.password)
        search_password.send_keys(Keys.RETURN)

        print("Logging successfully to JCB page")
        # time.sleep(2)

    def search_veritrans_page(self):
        # set the datas

        driver = self.driver

        driver.find_element_by_xpath("//li[@id='menu_search_b1']//a[contains(text(),'Order Search')]").click()

        driver.find_element_by_xpath("//ul[@id='payment_type']//label[1]//input[1]").click()

        select_date = Select(driver.find_element_by_id("default_re_transaction"))
        select_date.select_by_visible_text("Today")

        # get date forward
        i = datetime.datetime.now()

        # print ("dd/mm/yyyy format =  %s/%s/%s" % ((i.day+7), i.month, i.year))

        # type date into box
        search_end_date = driver.find_element_by_id("default_def_ddmmyy2")
        search_end_date.clear()
        search_end_date.send_keys(str(i.year) + '/' + str(i.month) + '/' + str((i.day + 7)))
        search_end_date.send_keys(u'\ue007')

        select_fetch = Select(driver.find_element_by_id("default_result_detail"))
        select_fetch.select_by_visible_text("2,000")

        # unselect elements
        driver.find_element_by_xpath("//div[@class='col-5']//input[2]").click()

        # select Capture and Cancel(Sales)
        driver.find_element_by_xpath("//div[@id='card_panel']//ul[2]//li[1]//label[1]//input[1]").click()
        driver.find_element_by_xpath("//div[@id='card_panel']//ul[2]//li[2]//label[1]//input[1]").click()

        # card company code - 02
        card_code = driver.find_element_by_id("card_company_card_code")
        card_code.send_keys("02")

        # final search
        driver.find_element_by_xpath("//input[@name='commit']").click()
        # driver.implicitly_wait(10)
        time.sleep(2)

        # auto downloading csv
        driver.find_element_by_xpath("//ul[@class='search_csv search_top_csv']//li[1]//form[1]").click()

        print("CSV successfully downloaded")

    def search_JCB_page(self):
        driver = self.driver

        driver.find_element_by_xpath("//a[@id='j_idt61']").click()
        driver.find_element_by_xpath("//a[@id='j_idt87_2:j_idt93']").click()

        print("PDF successfully downloaded ")

    # --- Post - Condition ---
    def tearDown(self):
        # to close the browser
        self.driver.close()
        print("browser successfully closed")

    # method to get the downloaded file name
    def getDownLoadedFileName(self, waitTime):
        driver = self.driver
        driver.execute_script("window.open()")
        # switch to new tab
        driver.switch_to.window(driver.window_handles[-1])
        # navigate to chrome downloads
        driver.get('chrome://downloads')
        # define the endTime
        endTime = time.time() + waitTime
        while True:
            try:
                # get downloaded percentage
                downloadPercentage = driver.execute_script(
                    "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList "
                    "downloads-item').shadowRoot.querySelector('#progress').value")
                # check if downloadPercentage is 100 (otherwise the script will keep waiting)
                if downloadPercentage == 100:
                    # return the file name once the download is completed
                    return driver.execute_script(
                        "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList "
                        "downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
            except:
                pass
            time.sleep(1)
            if time.time() > endTime:
                break

    def send_excel_to_powerautomate(self, file):
        # file = r"B100013100014503000145_20200717045252182.csv"  # excel file that you want to upload to Power Automate
        binary_upload_url = "https://prod-47.westus.logic.azure.com:443/workflows/8322569f0b6f4a689b442e5d0a521f8e" \
                            "/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1" \
                            ".0&sig=2wGKU63ZQThNTGt1Uo9xaTVg58jDeuQ-MCn3an_R_CU"  # url that you copy/paste form the
        # HTTP module in PowerAutomate
        data = open(file, 'rb').read()
        res = requests.post(binary_upload_url,
                            data=data,
                            headers={'Content-Type': 'application/octet-stream'})
        print("Power Automate status code: " + str(res.status_code))
        print("CSV successfully uploaded to Sharepoint")

    def rename_file(self, newname, folder_of_download, time_to_wait=60):
        time_counter = 0
        filename = max([f for f in os.listdir(folder_of_download)],
                       key=lambda xa: os.path.getctime(os.path.join(folder_of_download, xa)))
        while '.part' in filename:
            time.sleep(1)
            time_counter += 1
            if time_counter > time_to_wait:
                raise Exception('Waited too long for file to download')
        filename = max([f for f in os.listdir(folder_of_download)],
                       key=lambda xa: os.path.getctime(os.path.join(folder_of_download, xa)))
        os.rename(os.path.join(folder_of_download, filename), os.path.join(folder_of_download, newname))

        print("File successfully renamed")

    def delete_all_files(self, path):
        files = glob.glob(os.path.join(path))
        for file in files:
            os.remove(file)

        print("All files successfully deleted from directory")

    def read_csv(self):
        data = pandas.read_csv(r'C:\Users\guzulesc\PycharmProjects\JCB\jcb.csv')
        df = pandas.DataFrame(data, columns=['URL', 'Merchant ID', 'User ID', 'Password'])

        urls = df['URL']
        merchants = df['Merchant ID']
        users = df['User ID']
        passwords = df['Password']

        print("CSV successfully read")

        return urls, merchants, users, passwords


# if __name__ == "__main__":
    # #constructor with random parameters
    # temp = homePage("1", "2", "3", "4")
    # urls, merchants, users, passwords = temp.read_csv()
    #
    # list = []
    # for c in range(0, len(urls)):
    #     list.append(homePage(str(urls[c]), str(merchants[c]), str(users[c]), str(passwords[c])))
    #
    # # veritrans
    # list[2].get_chrome_driver()
    # list[2].test_load_home_page()
    # list[2].login_veritrans()
    # list[2].search_veritrans_page()
    # time.sleep(1)
    #
    # # JCB
    # # list[1].get_chrome_driver()
    # # list[1].test_load_home_page()
    # # list[1].login_JCB()
    # # list[1].search_JCB_page()
    #
    # # rename file
    # list[2].rename_file("JCB_2020.csv", r"C:\DownloadFiles\\", 60)
    #
    # # get name of the file
    # # entries = os.listdir('C:\DownloadFiles\\')
    # # print(entries[0])
    #
    # # get filename from chrome downloads tab
    # # latestDownloadedFileName = list[2].getDownLoadedFileName(5)  # waiting 5 seconds to complete the download
    # # print(latestDownloadedFileName)
    #
    # # upload to sharepoint
    # list[2].send_excel_to_powerautomate(r'C:\DownloadFiles\JCB_2020.CSV')
    #
    # # delete all files
    # list[2].delete_all_files(r'C:\DownloadFiles\*')
    #
    # # close driver
    # list[2].tearDown()
