import time
from selenium import webdriver
import getpass
import xlrd
from selenium.webdriver.chrome.options import Options

# ---------------------------------OPEN EXCEL-----------------------

loc = "tt1.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)


# ------------------Replacing New Lines-------------------
fin = open("ip.txt", "rt", encoding="utf-8")
fout = open("op.txt", "wt", encoding="utf-8")
for i in fin:
    fout.write(i.replace('\n', "%0a"))
    print("replaced")


fin.close()
fout.close()

# -------------MENU----------------------

fin = open("op.txt", "rt", encoding="utf-8")
print("Select Your Option ?")
select_type = int(input("1. Only Messages \n2. Image With Caption\n3. Media \n4. Media and Message Separately \nEnter :"))

if select_type == 1:
    msg = fin.read()
elif select_type == 3:
    file_path = input("Enter Media Path :")
else:
    msg = fin.read()
    file_path = input("Enter Media Path :")
fin.close()
# -------------------Settings--------------------
options = Options()
options.add_argument(f"--user-data-dir=C:/Users/{getpass.getuser()}/AppData/Local/Google/Chrome/User Data/Default")
options.add_argument("--profile-directory=Default")
options.headless = False
options.add_argument("--window-size=1920,1080")
options.add_argument('--ignore-certificate-errors')
options.add_argument('--allow-running-insecure-content')
options.add_argument("--disable-extensions")
options.add_argument("--proxy-server='direct://'")
options.add_argument("--proxy-bypass-list=*")
options.add_argument("--start-maximized")
options.add_argument('--disable-gpu/')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--no-sandbox')
options.add_argument("--log-level=3")
options.add_argument('--hide-scrollbars')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)

browser = webdriver.Chrome(options=options)
browser.get('https://web.whatsapp.com/')
browser.maximize_window()
time.sleep(10)
# val = select_type

# --------------------Function for message-----------------------


def message():

    print("Running......")
    for i in range(sheet.nrows):
        k = sheet.cell_value(i, 0)
        j = int(k)
        browser.get(f'https://api.whatsapp.com/send?phone={j}&text={msg}')
        time.sleep(2)
        browser.find_element_by_xpath("//*[@id='action-button']").click()
        time.sleep(1)
        browser.find_element_by_xpath("//*[@id='fallback_block']/div/div/a").click()
        time.sleep(2)
        try:
            if browser.find_element_by_xpath('//div[@data-animate-modal-popup="true"]'):
                browser.find_element_by_xpath('//div[@role="button"]').click()
                print(f"Invalid Number : {j}" + '\n')
                pass

        except Exception as e:
            time.sleep(2)
            browser.find_element_by_xpath(
                "/html/body/div/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div[3]/button").click()
            time.sleep(2)
            print(f"Successfully Send Msg : {j}" + '\n')
            time.sleep(2)

    print("Finish Task... \n")
    browser.quit()


# ---------------------------MEDIA--------------------------------


def media():
    # file_path = input("Enter Image Path :")
    print("Running......")
    for i in range(sheet.nrows):
        k = sheet.cell_value(i, 0)
        j = int(k)
        browser.get(f'https://api.whatsapp.com/send?phone={j}')
        time.sleep(2)
        browser.find_element_by_xpath("//*[@id='action-button']").click()
        time.sleep(1)
        browser.find_element_by_xpath("//*[@id='fallback_block']/div/div/a").click()
        time.sleep(5)
        try:
            if browser.find_element_by_xpath('//div[@data-animate-modal-popup="true"]'):
                browser.find_element_by_xpath('//div[@role="button"]').click()
                print(f"Invalid Number : {j}" + '\n')
                pass

        except Exception as e:
            time.sleep(2)
            attachment_section = browser.find_element_by_xpath('//div[@title = "Attach"]')
            attachment_section.click()
            image_box = browser.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
            image_box.send_keys(file_path)
            time.sleep(3)
            send_button = browser.find_element_by_xpath('//span[@data-icon="send"]')
            send_button.click()
            time.sleep(2)
            print(f"Successfully Send Msg : {j}" + '\n')
            time.sleep(5)
    print("Finish Task... \n")
    browser.quit()


# ---------------------------------Image with Caption---------------------------------


def img_with_txt():

    # file_path = input("Enter Image Path :")
    print("Running......")
    for i in range(sheet.nrows):
        k = sheet.cell_value(i, 0)
        j = int(k)
        browser.get(f'https://api.whatsapp.com/send?phone={j}&text={msg}')
        time.sleep(2)
        browser.find_element_by_xpath("//*[@id='action-button']").click()
        time.sleep(1)
        browser.find_element_by_xpath("//*[@id='fallback_block']/div/div/a").click()
        time.sleep(5)
        try:
            if browser.find_element_by_xpath('//div[@data-animate-modal-popup="true"]'):
                browser.find_element_by_xpath('//div[@role="button"]').click()
                print(f"Invalid Number : {j}" + '\n')
                pass

        except Exception as e:
            time.sleep(2)
            attachment_section = browser.find_element_by_xpath('//div[@title = "Attach"]')
            attachment_section.click()
            image_box = browser.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
            image_box.send_keys(file_path)
            time.sleep(3)
            send_button = browser.find_element_by_xpath('//span[@data-icon="send"]')
            send_button.click()
            time.sleep(2)
            print(f"Successfully Send Msg : {j}" + '\n')
            time.sleep(5)
    print("Finish Task... \n")
    browser.quit()

    # ----------------------MEDIA AND MESSAGE SEPARATELY--------------------------------


def two_files():
    # msg = input("Enter Messages : ")
    # file_path = input("Enter Image Path :")
    for i in range(sheet.nrows):
        k = sheet.cell_value(i, 0)
        j = int(k)
        browser.get(f'https://api.whatsapp.com/send?phone={j}')
        time.sleep(2)
        browser.find_element_by_xpath("//*[@id='action-button']").click()
        time.sleep(1)
        browser.find_element_by_xpath("//*[@id='fallback_block']/div/div/a").click()
        time.sleep(5)
        try:
            if browser.find_element_by_xpath('//div[@data-animate-modal-popup="true"]'):
                browser.find_element_by_xpath('//div[@role="button"]').click()
                print(f"Invalid Number : {j}" + '\n')
                pass

        except Exception as e:
            time.sleep(2)
            attachment_section = browser.find_element_by_xpath('//div[@title = "Attach"]')
            attachment_section.click()
            image_box = browser.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
            image_box.send_keys(file_path)
            time.sleep(3)
            send_button = browser.find_element_by_xpath('//span[@data-icon="send"]')
            send_button.click()
            time.sleep(2)
            print(f"Successfully Send Msg : {j}" + '\n')
            time.sleep(5)
        browser.get(f'https://api.whatsapp.com/send?phone={j}&text={msg}')
        time.sleep(2)
        browser.find_element_by_xpath("//*[@id='action-button']").click()
        time.sleep(1)
        browser.find_element_by_xpath("//*[@id='fallback_block']/div/div/a").click()
        time.sleep(2)
        try:
            if browser.find_element_by_xpath('//div[@data-animate-modal-popup="true"]'):
                browser.find_element_by_xpath('//div[@role="button"]').click()
                print(f"Invalid Number : {j}" + '\n')
                pass

        except Exception as e:
            time.sleep(2)
            browser.find_element_by_xpath(
                "/html/body/div/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div[3]/button").click()
            time.sleep(2)
            print(f"Successfully Send Msg : {j}" + '\n')
            time.sleep(2)
            print(f"Successfully Send Msg : {j}" + '\n')
            time.sleep(5)
    print("Finish Task... \n")
    browser.quit()

# -----------------------MAIN--------------------------------------------


if select_type == 1:
    message()
elif select_type == 2:
    img_with_txt()
elif select_type == 3:
    media()
elif select_type == 4:
    two_files()
else:
    print("Please enter a valid choice")
