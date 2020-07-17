import time
from m import homePage

if __name__ == '__main__':
    # constructor with random parameters
    temp = homePage("1", "2", "3", "4")
    urls, merchants, users, passwords = temp.read_csv()

    list = []
    for c in range(0, len(urls)):
        list.append(homePage(str(urls[c]), str(merchants[c]), str(users[c]), str(passwords[c])))

    # veritrans
    list[2].get_chrome_driver()
    list[2].test_load_home_page()
    list[2].login_veritrans()
    list[2].search_veritrans_page()
    time.sleep(1)

    # JCB
    # list[1].get_chrome_driver()
    # list[1].test_load_home_page()
    # list[1].login_JCB()
    # list[1].search_JCB_page()

    # rename file
    list[2].rename_file("JCB_2020.csv", r"C:\DownloadFiles\\", 60)

    # get name of the file
    # entries = os.listdir('C:\DownloadFiles\\')
    # print(entries[0])

    # get filename from chrome downloads tab
    # latestDownloadedFileName = list[2].getDownLoadedFileName(5)  # waiting 5 seconds to complete the download
    # print(latestDownloadedFileName)

    # upload to sharepoint
    list[2].send_excel_to_powerautomate(r'C:\DownloadFiles\JCB_2020.CSV')

    # delete all files
    list[2].delete_all_files(r'C:\DownloadFiles\*')

    # close driver
    list[2].tearDown()



