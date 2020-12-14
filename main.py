import autopy
import datetime
import time
from m import homePage

if __name__ == '__main__':
    # constructor with random parameters
    temp = homePage("1", "2", "3", "4", "5")
    tools, urls, merchants, users, passwords = temp.read_csv()

    portal = []
    for c in range(0, len(urls)):
        portal.append(homePage(str(tools[c]), str(urls[c]), str(merchants[c]), str(users[c]), str(passwords[c])))

    card_issuers = {
        "JCB": "02",
        "ORIENT CORPORATION": "10",
        "Toyota Finance": "31",
        "Sumitomo Mitsui Card Company": "04",
        "Mitsubishi UFJ NICOS": "05"
    }

    power_automate_urls = {
        "csv": "https://prod-47.westus.logic.azure.com:443/workflows/8322569f0b6f4a689b442e5d0a521f8e/triggers/manual"
               "/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig"
               "=2wGKU63ZQThNTGt1Uo9xaTVg58jDeuQ-MCn3an_R_CU",
        "xlsx": "https://prod-15.westus.logic.azure.com:443/workflows/86cf171932764f569e8a35d61c6d4f75/triggers"
                 "/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig"
                 "=-1GtsJduWC1bGNpl_lg2auFixm1iSTeZqRmKqyowJ2c",
        "pdf": "https://prod-67.westus.logic.azure.com:443/workflows/ccc47793ee3b41a5b7651d5efc5eabca" \
                            "/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1" \
                            ".0&sig=MFC35F2EIAMkFxureOEpWwmscfwZgEoKv1DUnkSrg00"
    }

    currentDay = datetime.datetime.now().day
    currentMonth = datetime.datetime.now().month
    currentYear = datetime.datetime.now().year

    csv_key = list(power_automate_urls.keys())[0]
    xls_key = list(power_automate_urls.keys())[1]
    pdf_key = list(power_automate_urls.keys())[2]

    jcb_key = list(card_issuers.keys())[0]
    orient_key = list(card_issuers.keys())[1]
    toyota_key = list(card_issuers.keys())[2]
    sumitomo_key = list(card_issuers.keys())[3]
    mitsubishi_key = list(card_issuers.keys())[4]

    download_path = "C:\DownloadFiles"
    SP_pdf_filename = "jcb.pdf"
    SP_csv_filename = "JCB_2020.csv"
    SP_xls_filename = "JCB_2020.xlsx"

    previousMonth = currentMonth

    if currentMonth == 1:
        previousMonth = 12

    # move mouse cursor out of the portal page for not causing any inconvenience
    autopy.mouse.move(700, 0)

    # if str(currentDay) == '10':
    if str(currentDay) == '10':

        # JCB - step 1
        portal[0].get_chrome_driver()
        portal[0].test_load_home_page()
        portal[0].login_JCB()
        jcb = portal[0].search_JCB_page()
        portal[0].rename(jcb[0], SP_pdf_filename, download_path)
        portal[0].send_to_powerautomate(f'{download_path}\{SP_pdf_filename}', power_automate_urls.get(pdf_key), pdf_key)
        portal[0].move_to_sharepoint()
        portal[0].rename_from_sharepoint(SP_pdf_filename, SP_pdf_filename[:-4], f"JCB_extractDay_{currentDay}_{jcb[0][:-4]}")
        portal[0].move_to_SP_folder(f"JCB_extractDay_{currentDay}_{jcb[0]}", currentYear, currentMonth)
        portal[0].delete_all_files(f'{download_path}\*')
        portal[0].close_driver()
        portal[0].tearDown()

    # if str(currentDay) == '5' or str(currentDay) == '20':
    if str(currentDay) == '5' or str(currentDay) == '10':

        # Orient: Orico - step 1.2.1
        portal[1].get_chrome_driver()
        portal[1].test_load_home_page()
        portal[1].login_plaza()
        plaza = portal[1].search_plaza(currentDay)
        portal[1].rename(plaza[0], SP_csv_filename, download_path)
        portal[1].send_to_powerautomate(f'{download_path}\{SP_csv_filename}', power_automate_urls.get(csv_key), csv_key)
        portal[1].move_to_sharepoint()
        portal[1].rename_from_sharepoint(SP_csv_filename, SP_csv_filename[:-4], f"B-PLAZA_extractDay_{currentDay}_{plaza[0][:-4]}")
        portal[1].delete_all_files(f'{download_path}\{SP_csv_filename}')

        portal[1].rename(plaza[1], SP_csv_filename, download_path)
        portal[1].send_to_powerautomate(f'{download_path}\{SP_csv_filename}', power_automate_urls.get(csv_key), csv_key)
        portal[1].rename_from_sharepoint(SP_csv_filename, SP_csv_filename[:-4], f"B-PLAZA_extractDay_{currentDay}_{plaza[1][:-4]}")
        portal[1].move_to_SP_folder(f"B-PLAZA_extractDay_{currentDay}_{plaza[0]}", currentYear, currentMonth)
        portal[1].move_to_SP_folder(f"B-PLAZA_extractDay_{currentDay}_{plaza[1]}", currentYear, currentMonth)
        portal[1].delete_all_files(f'{download_path}\*')
        portal[1].close_driver()
        portal[1].tearDown()


    # if str(currentDay) == '10' or str(currentDay) == '25':
    if str(currentDay) == '10' or str(currentDay) == '9':

        # Orient: B-plaza - step 1.2.2
        portal[2].get_chrome_driver()
        portal[2].test_load_home_page()
        portal[2].login_orico()
        files_names = portal[2].search_orico(currentYear, currentMonth)
        portal[2].rename(files_names[1], SP_pdf_filename, download_path)
        portal[2].rename(files_names[0], SP_csv_filename, download_path)
        portal[2].send_to_powerautomate(f'{download_path}\{SP_pdf_filename}', power_automate_urls.get(pdf_key), pdf_key)
        portal[2].send_to_powerautomate(f'{download_path}\{SP_csv_filename}', power_automate_urls.get(csv_key), csv_key)
        portal[2].move_to_sharepoint()
        portal[2].rename_from_sharepoint(SP_csv_filename, SP_csv_filename[:-4], f"OricoPayment_extractDay_{currentDay}_{files_names[0][:-4]}")
        portal[2].rename_from_sharepoint(SP_pdf_filename, SP_pdf_filename[:-4], f"OricoPayment_extractDay_{currentDay}_{files_names[1][:-4].lower()}")
        portal[2].move_to_SP_folder(f"OricoPayment_extractDay_{currentDay}_{files_names[0]}", currentYear, currentMonth)
        portal[2].move_to_SP_folder(f"OricoPayment_extractDay_{currentDay}_{files_names[1][:-4].lower() + files_names[1][-4:].lower()}", currentYear, currentMonth)
        portal[2].delete_all_files(f'{download_path}\*')
        portal[2].close_driver()
        portal[2].tearDown()

    # # if str(currentDay) == '16':
    # if str(currentDay) == '9':
    #
    #     # VERITRANS
    #     portal[3].get_chrome_driver()
    #     portal[3].test_load_home_page()
    #     portal[3].login_veritrans()
    #
    #     # if currentMonth == 1 or currentMonth == 2:
    #     if currentMonth == 1:
    #         d = portal[3].workdays(datetime.datetime(currentYear, previousMonth, 16),
    #                                datetime.datetime((currentYear + 1), currentMonth, 15))
    #     else:
    #         d = portal[3].workdays(datetime.datetime(currentYear, (currentMonth-1), 16),
    #                                datetime.datetime(currentYear, currentMonth, 15))
    #
    #     for credit_card_issuer, code in card_issuers.items():
    #
    #         for days in d:
    #             portal[3].search_veritrans_page(days.strftime("%Y/%m/%d"), code)
    #             print(days.strftime("%Y/%m/%d"))
    #             if days == d[-1]:
    #                 filenames = portal[3].get_file_name()
    #
    #         portal[3].merge_CSVs(filenames)
    #
    #         portal[3].send_to_powerautomate(f'{download_path}\{SP_csv_filename}', power_automate_urls.get(csv_key),
    #                                         csv_key)
    #         portal[3].move_to_sharepoint()
    #         portal[3].rename_from_sharepoint(SP_csv_filename, SP_csv_filename[:-4], credit_card_issuer)
    #         portal[3].move_to_SP_folder(credit_card_issuer + ".csv", currentYear, currentMonth)
    #         portal[3].delete_all_files(f'{download_path}\*')
    #         portal[3].close_driver()
    #
    #     portal[3].tearDown()