from selenium import webdriver
# from get_gecko_driver import GetGeckoDriver
import winsound
import time
import openpyxl
import logging




filename = f"jshshr.xlsx"
# logging.basicConfig(filename=f"{filename}.txt", level=logging.ERROR + logging.DEBUG + logging.FATAL + logging.CRITICAL + logging.WARNING + logging.INFO)


# d = GetGeckoDriver()
# d.install()

driver = webdriver.Firefox()
url = "https://doctor.medhub.uz/my-crews?crew_id=16131"


book = openpyxl.load_workbook('./jshshr_.xlsx')
data_sheet = book.active


def do_example(line):
    cancel = False
    driver.get(url)
    driver.maximize_window()
    data = data_sheet[f'a{line}'].value

    while True:
        try:
            bemorni_qoshish_tugamsi = driver.find_element('xpath', '/html/body/div/div[1]/div[3]/div/div[2]/div[1]/button')
            bemorni_qoshish_tugamsi.click()
            print("[LOG]: \"Bemor qo'shish\" bosildi")
            break
        except Exception as err:
            print("[ERROR]: \"Bemor qo'shish\" Bosilmadi")
        time.sleep(1)
    time.sleep(0.5)

    counter = 0
    refresh = False
    while True:
        if counter > 5:
            refresh = True
            break
        else:
            try:
                hujjat_turi_select = driver.find_element('xpath', '//*[@id="headlessui-listbox-button-5"]')
                hujjat_turi_select.click()
                print("[LOG]: \"Hujjat turi\" tanlanyabdi")
                break
            except Exception as err:
                print("[ERROR]: \"Hujjat turi\"", err)
            time.sleep(0.5)
        counter += 1
    if refresh:
        return

    while True:
        try:
            jshshir_option = driver.find_element('xpath', '//*[@id="headlessui-listbox-option-8"]')
            jshshir_option.click()
            print("[LOG]: \"JSHSHIR\" tanlandi")
            break
        except Exception as err:
            print("[ERROR]: \"JSHSHIR\"", err)

    while True:
        while True:
            try:
                jshshir_input = driver.find_element('xpath', '/html/body/div[2]/div/div/div/div[2]/div/div/div/div[1]/div/form/div[2]/label/input')
                for i in data:
                    jshshir_input.send_keys(i)
                    time.sleep(0.1)
                time.sleep(1)
                print("[LOG]: \"JSHSHIR\" yozildi")
                break
            except Exception as err:
                print("[ERROR]: \"JSHSHIR\"", err)
            time.sleep(0.5)
        
        try:
            wrong_check = driver.find_element('xpath', '/html/body/div[2]/div/div/div/div[2]/div/div/div/div[1]/div/form/div[2]/p')
            if "Notoʻgʻri JSHSHIR" in wrong_check.text:
                jshshir_input = driver.find_element('xpath', '/html/body/div[2]/div/div/div/div[2]/div/div/div/div[1]/div/form/div[2]/label/input')
                jshshir_input.clear()
                continue
            else:
                break
        except:
            break
 
    while True:
        try:
            topish_button = driver.find_element("xpath", "/html/body/div[2]/div/div/div/div[2]/div/div/div/div[1]/div/button")
            topish_button.click()
            time.sleep(2)
            try:
                elem = driver.find_element("xpath", "/html/body/div[2]/div/div/div/div[2]/div/div/div/div[1]/div/span")
                if elem.text == "Bemor topilmadi! Notoʼgʼri maʼlumotlar yoki server xatosi":
                    cancel = True
                    break
            except:
                pass
            print("[LOG]: \"Topish\" bosildi")
            break
        except Exception as err:
            print("[ERROR]: \"Topish\"", err)
        time.sleep(0.5)

    if cancel:
        data_sheet[f'a{line}'].value = data
        data_sheet[f'b{line}'].value = "Malumotlar bazasi ishlamadi"
        return

    while True:
        try:
            qoshish_tugmasi = driver.find_element('xpath', '/html/body/div[2]/div/div/div/div[2]/div/div/div/div[2]/button[2]')
            qoshish_tugmasi.click()
            print("[LOG]: \"Qo'shish\" bosildi")
            while 'cursor-wait' in qoshish_tugmasi.get_attribute('class'):
                time.sleep(0.5)
            break
        except Exception as err:
            try:
                bemorni_qoshish_tugamsi = driver.find_element('xpath', '/html/body/div/div[1]/div[3]/div/div[2]/div[1]/button')
                bemorni_qoshish_tugamsi.click()
                print("[LOG]: \"Bemor qo'shish\" bosildi")
            except Exception as err:
                print("[ERROR]: \"Bemor qo'shish\" Bosilmadi")
            print("[ERROR]: \"Qo'shish\"", err)
        time.sleep(0.5)

    data_sheet[f'a{line}'].value = data
    while True:
        try:
            errors = driver.find_element('xpath', '/html/body/div[1]/div[2]')
            data_sheet[f'a{line}'].value = data
            if "This profile is already" in errors.text:
                data_sheet[f'b{line}'].value = "Qo'shilgan"
            print("[LOG]: Tugatildi")
            break
        except Exception as err:
            print("[ERROR]: Tugash", err, "tomonidan chopildi")

    book.save(filename)

pk = int(input('Kelgan soni: '))
driver.get(url)

input("Start")

try:
    while True:
        do_example(pk)

        if data_sheet[f'a{pk + 1}'].value != None:
            pk += 1
        else:
            pk = 1
except:
    pass


driver.delete_all_cookies()
driver.quit()
