import os
import csv
import glob
import time
import xlwt
import json
import pandas as pd
import calendar
import datetime
import requests
from selenium import webdriver
import win32com.client as win32
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager



class homePage(object):
    PATH_TO_CHROME_DRIVER = r"C:\chromedriver\chromedriver.exe"

    def __init__(self, tools, base_url, merchant, username, password):
        self.tools = tools
        self.base_url = base_url
        self.merchant = merchant
        self.username = username
        self.password = password

    def get_chrome_driver(self):
        chrome_options = Options()
        prefs = {
            "download.default_directory": r"C:\DownloadFiles",
            "safebrowsing.enabled": "false"}
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("user-data-dir=selenium")

        self.driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)

        # browser should be loaded in maximized window
        self.driver.maximize_window()

        # driver should wait implicitly for a given duration, for the element under consideration to load.
        self.driver.implicitly_wait(10)  # 10 is in seconds

        print("Chrome browser successfully opened")

    # --- Steps ---
    def test_load_home_page(self):
        # to initialize a variable to hold reference of webdriver instance being passed to the function as a reference.
        driver = self.driver
        # to load a given URL in browser window
        driver.get(self.base_url)

        print(f"{self.tools} page successfully loaded")

    def download_wait(self, path_to_downloads):
        seconds = 0
        dl_wait = True
        while dl_wait and seconds < 20:
            time.sleep(1)
            dl_wait = False
            for fname in os.listdir(path_to_downloads):
                if fname.endswith('.crdownload'):
                    dl_wait = True
            seconds += 1
        return seconds

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

        print(f"Logging successfully to {self.tools} portal")

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

        print(f"Logging successfully to {self.tools} portal")
        # time.sleep(2)

    def login_orico(self):
        driver = self.driver
        print(driver.current_url)

        ID = driver.find_element_by_id("ID_TXT").send_keys(self.username)
        PW = driver.find_element_by_id("PASS_TXT").send_keys(self.password)

        driver.find_element_by_id("OK_BTN").click()

        print(f"Logging successfully to {self.tools} portal")

    def login_plaza(self):
        driver = self.driver
        print(driver.current_url)


        ID = driver.find_element_by_id("LoginBplazaUserId").send_keys(self.username)
        PW = driver.find_element_by_id("LoginBplazaPassword").send_keys(self.password)

        driver.find_element_by_id("LoginBtn").click()

        print(f"Logging successfully to {self.tools} portal")

    def search_orico(self, currentYear, currentMonth):
        driver = self.driver
        # currentMonth = datetime.datetime.now().month
        # currentYear = datetime.datetime.now().year
        #
        # print(f"Current month = {currentMonth}")
        # print(f"Current year = {currentYear}")

        year = driver.find_element_by_id("YEAR_TXT")
        year.clear()
        year.send_keys(currentYear)

        month = driver.find_element_by_id("MONTH_CMB")
        select = Select(month)
        select.select_by_value(str(currentMonth))

        driver.find_element_by_id("HYOUJI_BTN").click()
        time_pdf = self.download_wait("C:\DownloadFiles")
        print(f'It took {time_pdf} seconds to fully download the pdf')

        driver.find_element_by_id("CSVSYUTURYOKU_btn").click()
        time_csv = self.download_wait("C:\DownloadFiles")
        print(f'It took {time_csv} seconds to fully download the csv')

        # get name of the csv downloaded from Orico Payment
        # files_names = os.listdir('C:\DownloadFiles\\')
        # print("Files names are = ", files_names)
        # time.sleep(1)

        return self.get_file_name()

    def search_plaza(self, currentDay):
        driver = self.driver
        driver.find_element_by_xpath(
            "//body/form[@name='KLLAAA0011']/div[@id='layout_page']/div[@id='layout_menu']/div[@id='bpz_menu']/ul/li["
            "2]/a[1]").click()

        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, "//li[2]//ul[1]//li[1]//a[1]")))
        actions = ActionChains(driver)
        actions.move_to_element(element).click(element).perform()


        driver.find_element_by_xpath("//a[@class='idLink_payStoreId']").click()

        dropdown = Select(driver.find_element_by_xpath("//select[@id='SiharaiMonthIndex']"))
        if str(currentDay) == '20':
            dropdown.select_by_value("0")
            driver.find_element_by_xpath("//tr//tr[1]//td[1]//a[1]").click()

        if str(currentDay) == '10':
            dropdown.select_by_value("1")
            driver.find_element_by_xpath("//tr[2]//td[1]//a[1]").click()

        driver.find_element_by_xpath("//a[@class='csvLinkSiharai']").click()

        time_csv1 = self.download_wait("C:\DownloadFiles")
        print(f'It took {time_csv1} seconds to fully download the first csv')

        driver.find_element_by_xpath("//a[@class='tabLink_meisai']").click()

        dropdown2 = Select(driver.find_element_by_xpath("//select[@id='MeisaiJyoukenUriageKbn']"))
        dropdown2.select_by_value("3")

        driver.find_element_by_xpath("//input[@id='SearchBtn']").click()
        driver.find_element_by_xpath("//a[@class='csvLink']").click()

        time_csv2 = self.download_wait("C:\DownloadFiles")
        print(f'It took {time_csv2} seconds to fully download the second csv')

        return self.get_file_name()

    def search_veritrans_page(self, the_date, company_code):
        driver = self.driver

        driver.find_element_by_xpath("//li[@id='menu_search_b1']//a[contains(text(),'Order Search')]").click()
        driver.find_element_by_xpath("//ul[@id='payment_type']//label[1]//input[1]").click()

        search_begin_date = driver.find_element_by_id("default_def_mmddyy1")
        search_begin_date.clear()
        search_begin_date.send_keys(the_date)
        search_begin_date.send_keys(u'\ue007')

        search_end_date = driver.find_element_by_id("default_def_ddmmyy2")
        search_end_date.clear()
        search_end_date.send_keys(the_date)
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
        card_code.send_keys(company_code)

        # final search
        driver.find_element_by_xpath("//input[@name='commit']").click()

        # auto downloading csv
        wait = WebDriverWait(driver, 20)
        element = wait.until(
            EC.presence_of_element_located((By.XPATH, "//ul[@class='search_csv search_top_csv']//li[1]//form[1]")))
        actions = ActionChains(driver)
        actions.move_to_element(element)
        time.sleep(1)
        driver.execute_script("window.scrollTo(0, 1000)")
        time.sleep(1)
        actions.click(element).perform()

        time_csv = self.download_wait("C:\DownloadFiles")
        print(f'It took {time_csv} seconds to fully download the pdf')

    def search_JCB_page(self):
        driver = self.driver

        driver.find_element_by_id("j_idt61").click()
        driver.find_element_by_xpath(
            "/html[1]/body[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/div[2]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[2]/p[1]/a[1]").click()

        time_csv = self.download_wait("C:\DownloadFiles")
        print(f'It took {time_csv} seconds to fully download the pdf')

        print("PDF successfully downloaded ")

        # get name of the csv downloaded from Orico Payment
        # jcb = os.listdir('C:\DownloadFiles\\')
        # print("Files names are = ", jcb)
        # time.sleep(1)

        return self.get_file_name()

    # --- Post - Condition ---
    def tearDown(self):
        # to close the browser
        time.sleep(2)
        self.driver.close()
        print("browser successfully closed")

    def get_file_name(self):
        file_name = os.listdir('C:\DownloadFiles\\')
        print("Files names are = ", file_name)
        time.sleep(1)

        return file_name

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

    def rename(self, filename, newname, folder_of_download):
        os.rename(os.path.join(folder_of_download, filename), os.path.join(folder_of_download, newname))
        print("File successfully renamed")

    def send_to_powerautomate(self, file, binary_upload_url, file_type):
        # file = r"B100013100014503000145_20200717045252182.csv"  # excel file that you want to upload to Power Automate
        # binary_upload_url = "https://prod-67.westus.logic.azure.com:443/workflows/ccc47793ee3b41a5b7651d5efc5eabca" \
        #                     "/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1" \
        #                     ".0&sig=MFC35F2EIAMkFxureOEpWwmscfwZgEoKv1DUnkSrg00"
        # # url that you copy/paste form the
        # HTTP module in PowerAutomate
        data = open(file, 'rb').read()
        res = requests.post(binary_upload_url,
                            data=data,
                            headers={'Content-Type': 'application/octet-stream'})
        print("Power Automate status code: " + str(res.status_code))
        print(f"{file_type} successfully uploaded to Sharepoint")

    def upload_text_file_to_powerautomate(self, site_address, file_location, file_name, target_path):
        """
            Description: uploads a text file to a SharePoint folder, whose path is specified using a PowerAutomate flow.
                Make sure the flow is turned on
            :site_address: SharePoint site address
            :target_path: Path to the folder in the SharePoint site where the file will be uploaded to
            :file_location: Local folder in the sending computer that is sending the file
            :file_name: Name of the file that will be uploaded
        """
        power_automate_url = "https://prod-45.westus.logic.azure.com:443/workflows/394b4e6c021a4c97a886a90ef2bc4aba/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=zl2BI26XhYx9cFf0oKh0NcVT-BBYLGfi0HKj5p-suyE"  # copy / paste this URL from PowerAutomate HTTPS component
        full_file_location = os.path.join(file_location, file_name)  # puts together the file location and file name
        with open(full_file_location, 'rb') as f:
            file_content = f.read()  # read the content of the file
            data = {
                'site_address': site_address,
                'target_path': target_path,
                'file_name': file_name,
                'file_content': file_content
            }
            data_json = json.dumps(data)  # converts the dictionary to a json format
            r = requests.post(power_automate_url,
                              json=data_json)  # uploads the json data to the PowerAutomate HTTP web service

    def delete_all_files(self, path):

        try:
            files = glob.glob(os.path.join(path))
            for file in files:
                os.remove(file)
        except:
            pass

        print("All files successfully deleted from directory")

    def read_csv(self):
        data = pd.read_csv(r'C:\Users\guzulesc\PycharmProjects\JCB\jcb.csv')
        df = pd.DataFrame(data, columns=['Tools', 'URL', 'Merchant ID', 'User ID', 'Password'])

        tools = df['Tools']
        urls = df['URL']
        merchants = df['Merchant ID']
        users = df['User ID']
        passwords = df['Password']

        print("CSV successfully read")

        return tools, urls, merchants, users, passwords

    # def open_new_tab(self):
    #     driver = self.driver
    #     driver.execute_script("window.open('');")
    #     driver.switch_to.window(driver.window_handles[1])

    def move_to_sharepoint(self):
        driver = self.driver
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[1])
        driver.get('https://hp.sharepoint.com/sites/CCRobotics/RocketAPJ%20Credit%20Card/Forms/AllItems.aspx')

        driver.implicitly_wait(5)
        if not driver.find_element_by_id('appRoot').text:
            driver.find_element_by_name('loginfmt').send_keys('sebastian-nicolae.guzulescu@hp.com')
            driver.find_element_by_name('loginfmt').send_keys(u'\ue007')
            time.sleep(10)
            driver.implicitly_wait(80)

            driver.find_element_by_class_name('middle')
            print("Found middle of the window with yes button")
            driver.find_element_by_class_name('btn').click()
        else:
            print("Already logged in!")
        # pickle.dump(self.driver.get_cookies() , open("cookies.pkl","wb"))
        # print("Typed in login email and click next")

    def rename_from_sharepoint(self, filename, filename_without_extension, newname):
        driver = self.driver
        driver.refresh()

        try:

            wait = WebDriverWait(driver, 30)
            # element = driver.find_element_by_xpath('//button[text()="JCB_2020.csv"]')
            element = wait.until(EC.visibility_of_element_located((By.XPATH, f'//button[text()="{filename}"]')))
            # time.sleep(1)
            actions = ActionChains(driver)
            actions.move_to_element(element).context_click(element).perform()
            # ActionChains(driver).context_click(element).perform()
            time.sleep(1)

            # buton de apasat pe rename
            driver.find_element_by_xpath("//li[11]").click()

            driver.find_element_by_xpath(f"//input[@value='{filename_without_extension}']").send_keys(newname + Keys.RETURN)

            # pyautogui.typewrite("JCB from " + str(begin_date) + " to " + str(end_date))
            time.sleep(.5)
            # driver.find_element_by_xpath("//button[@class='ms-Button ms-Button--primary root-244']//span[@class='ms-Button-textContainer textContainer-107']").click()
            # time.sleep(.5)
            driver.find_element_by_xpath(
                "//div[@class='ms-OverflowSet ms-CommandBar-secondaryCommand secondarySet-116']//div[1]").click()
            time.sleep(.5)

        # # check if "csv_2020.csv" exists in sharepoint
        # elem = []
        # try:
        #     elem.append(driver.find_element_by_xpath('//button[text()="JCB_2020.csv"]'))
        #     if len(elem) > 0:
        #         actions.move_to_element(elem[0]).context_click(elem[0]).perform()
        #         time.sleep(1)
        #         driver.find_element_by_xpath("//li[8]").click()
        #         driver.find_element_by_xpath("/html[1]/body[1]/div[7]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[2]/div[1]/span[1]/button[1]").click()
        # except:
        #     pass
        except:
            pass
        # driver.switch_to.window(driver.window_handles[0])
        # driver.close()
        # time.sleep(.5)
        # driver.switch_to.window(driver.window_handles[0])

    def move_to_SP_folder(self, filename, currentYear, currentMonth):
        driver = self.driver

        try:

            wait = WebDriverWait(driver, 10)
            # element = driver.find_element_by_xpath('//button[text()="JCB_2020.csv"]')
            element = wait.until(EC.visibility_of_element_located((By.XPATH, f'//button[text()="{filename}"]')))
            # time.sleep(1)
            actions = ActionChains(driver)
            actions.move_to_element(element).context_click(element).perform()
            # ActionChains(driver).context_click(element).perform()
            time.sleep(1)

            # buton de apasat pe 'move to'
            driver.find_element_by_xpath("//li[13]").click()

            driver.find_element_by_xpath("//span[contains(text(),'Current Library')]").click()
            time.sleep(.7)

            for folder in driver.find_elements_by_xpath(f"//span[contains(text(),{currentYear})]"):
                folder.click()
                time.sleep(.7)

            for folder in driver.find_elements_by_xpath(f"//span[contains(text(),{currentMonth})]"):
                folder.click()
                time.sleep(.7)

            driver.find_element_by_xpath("//span[@class='od-Button-label']").click()

            time.sleep(2)

        except:
            pass

    def send_email_via_outlook(self):
        # outlook = win32.Dispatch('outlook.application')

        try:
            outlook = win32.GetActiveObject('Outlook.Application')
        except:
            outlook = win32.Dispatch('Outlook.Application')

        mail = outlook.CreateItem(0)
        # mail.To = 'sebastian-nicolae.guzulescu@hp.com;  alexandru-bogdan.stefan@hp.com; robert.rudberg@hp.com; ' \
        #           'cristina-andreea.melente@hp.com; georgiana.lichi1@hp.com '
        mail.To = 'sebastian-nicolae.guzulescu@hp.com'
        mail.Subject = 'APJ script has finished'
        # mail.Body = 'Message body'
        mail.HTMLBody = '<h2>All CSVs have been successfully uploaded to the sharepoint.</h2>'
        mail.Send()
        time.sleep(5)
        print("Mail successfully sent")

    def check_if_last_day_of_month(self, date):
        #  calendar.monthrange return a tuple (weekday of first day of the
        #  month, number
        #  of days in month)
        last_day_of_month = calendar.monthrange(date.year, date.month)[1]
        # here i check if date is last day of month
        print(f'Last day of {calendar.month_name[datetime.datetime.now().month]} is {last_day_of_month}')
        print(f'Today is {calendar.month_name[datetime.datetime.now().month]} {datetime.datetime.now().day}')
        print(f'Date that should be checked {datetime.date(date.year, date.month, 15)}')
        # if date == datetime.date(date.year, date.month, last_day_of_month):
        if date == datetime.date(date.year, date.month, 15):
            return True
        return False

    def download_CSVs_from_sharepoint(self, begin_date):
        driver = self.driver
        try:
            element = driver.find_element_by_xpath(f"//*[contains(@aria-label, '{begin_date}')]")
            time.sleep(.2)
            # driver.execute_script("arguments[0].scrollIntoView();", element)

            # element = driver.execute_script("arguments[0].scrollIntoView(true);", el)
            # time.sleep(1)

            # wait = WebDriverWait(driver, 10)
            # element = wait.until(EC.visibility_of_element_located((By.XPATH, f"//*[contains(@aria-label, '{begin_date}')]")))
            actions = ActionChains(driver)
            actions.move_to_element(element).context_click(element).perform()
            time.sleep(1)

            # buton de apasat pe download
            driver.find_element_by_xpath("//li[7]").click()
            time.sleep(1.5)

            driver.find_element_by_xpath("//button[@name='Delete']").click()
            time.sleep(1)

            # driver.find_element_by_xpath('//*[contains(@title, "Close")]').click()
            driver.find_element_by_xpath('//*[contains(@class, "ms-Dialog-action")]').click()
            time.sleep(.2)
            # LAST

        except:
            pass

    # def sort_descending_by_name(self):
    #     driver = self.driver
    #     driver.find_element_by_xpath(r"/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[2]/div[2]/div["
    #                                  r"2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div["
    #                                  r"1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[5]/div[1]").click()
    #     driver.find_element_by_xpath("//span[contains(text(),'Z to A')]").click()

    def close_driver(self):
        driver = self.driver
        driver.close()
        time.sleep(.5)
        driver.switch_to.window(driver.window_handles[0])

    def merge_CSVs(self):
        # path = r'C:\DownloadFiles'  # use your path
        # all_files = glob.glob(path + "/*.csv")
        #
        # li = []
        #
        # for filename in all_files:
        #     df = pd.read_csv(filename, index_col=None, header=0, encoding= 'unicode_escape')
        #     li.append(df)
        #
        # frame = pd.concat(li, axis=0, ignore_index=True, sort=False)
        # frame.to_csv("C:\DownloadFiles\JCB_2020.csv")

        path = r'C:\DownloadFiles'  # use your path
        file_extension = '.csv'
        all_filenames = [i for i in glob.glob(f"{path}/*{file_extension}")]
        combined_csv_data = pd.concat([pd.read_csv(f, encoding='unicode_escape') for f in all_filenames])
        combined_csv_data.to_csv('C:\DownloadFiles\JCB_2020.csv')  # Saving our combined csv data as a new file!

        print("CSVs were successfully concatenated!")

    def from_csv_to_excel(self, path):

        for csvfile in glob.glob(os.path.join(path, '*.csv')):
            wb = xlwt.Workbook()
            ws = wb.add_sheet('data')
            with open(csvfile, 'rt', encoding='unicode-escape') as f:
                reader = csv.reader(f)
                for r, row in enumerate(reader):
                    for c, val in enumerate(row):
                        ws.write(r, c, val)
            wb.save(csvfile[:-4] + '.xls')

    def rename_monthly_CSV_from_sharepoint(self, begin_date, end_date):
        driver = self.driver

        try:

            wait = WebDriverWait(driver, 10)
            # element = driver.find_element_by_xpath('//button[text()="JCB_2020.csv"]')
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//button[text()="JCB_2020.csv"]')))
            # time.sleep(1)
            actions = ActionChains(driver)
            actions.move_to_element(element).context_click(element).perform()
            # ActionChains(driver).context_click(element).perform()
            time.sleep(1)

            # buton de apasat pe rename
            driver.find_element_by_xpath("//li[11]").click()

            driver.find_element_by_xpath("//input[@value='JCB_2020']").send_keys(
                "MONTHLY csv from " + str(begin_date) + " to " + str(end_date) + Keys.RETURN)

            # pyautogui.typewrite("JCB from " + str(begin_date) + " to " + str(end_date))
            time.sleep(.5)
            # driver.find_element_by_xpath("//button[@class='ms-Button ms-Button--primary root-244']//span[@class='ms-Button-textContainer textContainer-107']").click()
            # time.sleep(.5)
            driver.find_element_by_xpath(
                "//div[@class='ms-OverflowSet ms-CommandBar-secondaryCommand secondarySet-116']//div[1]").click()
            time.sleep(.5)

        except:
            pass
        # driver.switch_to.window(driver.window_handles[0])
        driver.close()
        time.sleep(.5)
        driver.switch_to.window(driver.window_handles[0])

    def merge_CSVs(self, first_file_name):
        dir = os.chdir('C:\DownloadFiles\\')
        fout = open("JCB_2020.csv", "ab")
        # first file:
        for line in open(first_file_name[0], 'rb'):
            fout.write(line)

        print(first_file_name[0] + "has been added to single file" + '\n')
        print(first_file_name[0] + "has been deleted")
        del first_file_name[0]
        # now the rest:
        list = os.listdir(dir)
        number_files = len(list)
        for num in range(2, number_files):
            f = open(first_file_name[0], 'rb')
            f.__next__()  # skip the header
            for line in f:
                fout.write(line)
            print(first_file_name[0] + "has been added to single file" + '\n')
            print(first_file_name[0] + "has been deleted")
            del first_file_name[0]

            f.close()  # not really needed
        fout.close()

    def workdays(self, d, end):
        days = []
        while d.date() <= end.date():
            days.append(d)
            d += datetime.timedelta(days=1)
        return days


