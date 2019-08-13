from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import csv
import time
import os
from tkinter import *
from tkinter import messagebox
import datetime
import sys
import threading
import logging


def getLicence():
    date = datetime.datetime.now()
    nowDate = date.date()
    key = '2018-11-27'

    keyDate = datetime.datetime.strptime(key,"%Y-%m-%d").date()
    if(nowDate > keyDate):
        return False
    else:
        return True


def closeChecking():
    global thread_flag
    thread_flag = False

    try:
        fw.close()
    except:
        pass

    try:
        window.destroy()
    except:
        pass

    try:
        driver.quit()
    except:
        pass

    sys.exit()


def validation(codes, total, start, end, interval, reload_interval):
    global reload_interval_value

    if(interval == ''):
        interval = '0'
    try:
        startNumber = int(start)
    except:
        messagebox.showerror('ERROR!', 'Start number must be integer!.\nPlease try again.')
        return
    if(startNumber<1):
        messagebox.showerror('ERROR!', 'Start number must be bigger than 0!.\nPlease try again.')
        return
    try:
        endNumber = int(end)
    except:
        messagebox.showerror('ERROR!', 'End number must be integer!.\nPlease try again.')
        return
    if (endNumber > total):
        messagebox.showerror('ERROR!', 'End number must be less than '+str(total)+' .\nPlease try again.')
        return
    try:
        intervalNumber = int(interval)
    except:
        messagebox.showerror('ERROR!', 'Interval must be integer!.\nPlease try again.')
        return
    if (intervalNumber < 0):
        messagebox.showerror('ERROR!', 'End number must be bigger than 0.\nPlease try again.')
        return

    try:
        reload_interval_value = int(reload_interval)
    except:
        messagebox.showerror('ERROR!', 'Reload interval must be integer!.\nPlease try again.')
        return
    if (reload_interval_value < 0):
        messagebox.showerror('ERROR!', 'End number must be bigger than 0.\nPlease try again.')
        return

    ourThread = threading.Thread(target=doCheck, args=(codes, startNumber, endNumber, intervalNumber))
    ourThread.start()


def multiTab_closing():
    window_handles = driver.window_handles
    if len(window_handles) > 1:
        main_windowHandle = window_handles[0]

        for handle in window_handles[1:]:
            driver.switch_to.window(handle)
            driver.close()

        driver.switch_to.window(main_windowHandle)


def doCheck(codes, nStart, nEnd, nInterval):
    m = nStart
    reload_counter = 0
    for j in range(nStart, nEnd, 1):
        logger.info('Reload counter: %s' % reload_counter)
        refresh_attempt = 0

        if (not thread_flag):
            return

        multiTab_closing()

        if j < m:
            continue

        if reload_counter >= reload_interval_value:
            print('Refreshing by Reload interval.')
            logger.info('Refreshing by Reload interval.')
            try:
                driver.refresh()
                refresh_attempt += 1
                reload_counter = 0
            except TimeoutException:
                refresh_attempt = max_refresh_attempt

        while True:
            try:
                m, success = onCart(j, m, codes, nInterval)
                if not success:
                    reload_counter += 1
                else:
                    reload_counter = 0

            except Exception as e:
                logger.error(e)
                reload_counter = 0

                if (not thread_flag):
                    return

                if refresh_attempt >= max_refresh_attempt:
                    driver.quit()
                    print('Restart')
                    logger.info('Restart')
                    init_driver()
                    go_checkPage()
                    refresh_attempt = 0
                else:
                    print('Refreshing')
                    logger.info('Refreshing')
                    try:
                        driver.refresh()
                        refresh_attempt += 1
                    except TimeoutException:
                        refresh_attempt = max_refresh_attempt
            else:
                break

    fw.close()
    driver.close()
    messagebox.showinfo('Congratulations!', 'All codes have been checked successfully!')
    window.destroy()
    quit()
    exit()
    sys.exit(0)


def onCart(j, m, codes, nInterval):
    logger.info(str(codes[j-1]))

    openFlag = True
    while openFlag:
        try:
            promo_input = waitShort.until(EC.visibility_of_element_located((By.ID, 'promocode-input')))
            promo_input.send_keys(str(codes[j - 1]))
            openFlag = False
        except Exception as e:
            time.sleep(1)
            driver.find_element_by_xpath('//a[@href="#promoToggle"]').click()
            try:
                remove = waitShortShort.until(EC.visibility_of_element_located((By.CLASS_NAME, 'js-remove-promo')))
                remove.click()
            except:
                pass

            openFlag = True

    promo_btn = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'js-promo-button')))
    promo_btn.click()
    time.sleep(5)

    try:
        errorBox = driver.find_element_by_class_name('js-promo-error-alert')
        errorMsg = errorBox.text
        flag = False
    except:
        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[contains(@class, "art-sc-yourTotSavingsValue")]')))
            # driver.find_element_by_xpath('/*[contains(@class, "art-sc-yourTotSavingsValue")]')
            flag = True
        except Exception as e:
            flag = False

    if flag:
        driver.find_element_by_link_text('Add Promotional Code').click()
        time.sleep(3)
        remove = waitShort.until(EC.visibility_of_element_located((By.CLASS_NAME, 'js-remove-promo')))
        remove.click()

        strCode = str(codes[j - 1])
        print("Valid Code: %s" % strCode)
        logger.info("Valid Code: %s" % strCode)
        fw.write(strCode)

        return j + nInterval, True

    else:
        driver.find_element_by_name('promocode').clear()

        return j, False


def go_checkPage():
    attempt = 0
    while True:
        try:
            driver.get(item_url)
            try:
                element = wait.until(
                    EC.visibility_of_element_located((By.XPATH, '//img[contains(@alt, "Click to close chat invitation")]')))
                element.click()
            except:
                pass

            element = waitLong.until(EC.visibility_of_element_located((By.XPATH, '//button[contains(text(), "Add To Cart")]')))
            time.sleep(5)
            element.click()
            # driver.execute_script("arguments[0].click();", element.find_element_by_xpath('..'))

            element = waitLong.until(EC.visibility_of_element_located((By.XPATH, '//a[contains(text(), "View Cart")]')))
            time.sleep(5)
            element.click()
            # driver.execute_script("arguments[0].click();", element)

        except:
            if attempt >= max_refresh_attempt:
                time.sleep(30)
                print('Restarting')
                logger.info('Restarting')
                driver.quit()
                init_driver()
                attempt = 0
            else:
                print('Refresh')
                logger.info('Refresh')
                attempt += 1
        else:
            break


def mainForm():
    global fw
    global window
    global item_url

    if not os.path.isdir('data'):
        os.system('mkdir data')

    try:
        fr = open('config.csv')
    except:
        messagebox.showerror('ERROR!','config.csv not exist in directory to.\nplease put that file in there and try again!')
        return

    reader = csv.DictReader(fr)

    input_file_name = ''
    output_file_name = ''
    item_url = ''
    for row in reader:
        item_url = row['item_url']
        if not item_url.strip():
            print("No Item URL")
            sys.exit()

        input_file_name = row['input_file_name']
        output_file_name = row['output_file_name']

    fr.close()

    fw = open('data/'+output_file_name+'.txt', 'w')

    rd = open(input_file_name+'.txt')

    codes = rd.readlines()
    firstCode = codes[0]
    lastCode = codes[-1]
    length = len(codes)
    rd.close()

    window = Tk()

    window.title('Code Checker')
    window.geometry("400x500+1000+200")
    window.resizable(False, False)

    labelTotal = Label(window, text='Total Count: ' + str(length), fg='blue')
    labelTotal.pack()
    labelTotal.config(font=("Courier", 12))
    labelTotal.place(x=50, y=20)

    labelInput = Label(window, text='Input file: ' + input_file_name + '.txt', fg='blue')
    labelInput.pack()
    labelInput.config(font=("Courier", 12))
    labelInput.place(x=60, y=60)

    labelOutput = Label(window, text='Output file: ' + output_file_name + '.txt', fg='blue')
    labelOutput.pack()
    labelOutput.config(font=("Courier", 12))
    labelOutput.place(x=50, y=100)

    labelStartCode = Label(window,text='Start code: '+ str(firstCode))
    labelStartCode.pack()
    labelStartCode.config(font=("Courier", 12))
    labelStartCode.place(x=60, y=150)

    labelLastCode = Label(window, text='Last code: ' + str(lastCode))
    labelLastCode.pack()
    labelLastCode.config(font=("Courier", 12))
    labelLastCode.place(x=70, y=200)

    labelStart = Label(window, text='Start Number:')
    labelStart.pack()
    labelStart.config(font=("Courier", 14))
    labelStart.place(x=30, y=250)

    inputStart = Entry(window)
    inputStart.pack()
    inputStart.place(x=180, y=250, width=200)
    inputStart.config(font=("Courier", 16))

    labelEnd = Label(window, text='End Number:')
    labelEnd.pack()
    labelEnd.config(font=("Courier", 14))
    labelEnd.place(x=50, y=300)

    inputEnd = Entry(window)
    inputEnd.pack()
    inputEnd.place(x=180, y=300, width=200)
    inputEnd.config(font=("Courier", 16))

    labelInterval = Label(window, text='Interval:')
    labelInterval.pack()
    labelInterval.config(font=("Courier", 14))
    labelInterval.place(x=70, y=350)

    inputInterval = Entry(window)
    inputInterval.pack()
    inputInterval.place(x=180, y=350, width=200)
    inputInterval.config(font=("Courier", 16))

    label_reloadInterval = Label(window, text='Reload Interval:')
    label_reloadInterval.pack()
    label_reloadInterval.config(font=("Courier", 14))
    label_reloadInterval.place(x=0, y=400)

    reloadInterval = Entry(window)
    reloadInterval.pack()
    reloadInterval.place(x=180, y=400, width=200)
    reloadInterval.config(font=("Courier", 16))

    btnStart = Button(window, text='START', bg="orange",
                      command=lambda: validation(
                          codes, length, inputStart.get(), inputEnd.get(), inputInterval.get(), reloadInterval.get()))
    btnStart.pack()
    btnStart.place(x=70, y=450, width=100)
    btnStart.config(font=("Courier", 14))

    btnClose = Button(window, text='CLOSE', bg="skyblue",command=lambda: closeChecking())
    btnClose.pack()
    btnClose.place(x=200, y=450, width=100)
    btnClose.config(font=("Courier", 14))

    go_checkPage()
    # driver.get(item_url)

    window.mainloop()


def init_driver():
    global driver
    global wait
    global waitLong
    global item_url
    global waitLongLong
    global waitShort
    global waitShortShort

    os.environ['NO_PROXY'] = 'lowes.com'
    # capa = DesiredCapabilities.CHROME
    # capa["pageLoadStrategy"] = "none"

    option = webdriver.ChromeOptions()
    option.add_argument('--ignore-certificate-errors')
    option.add_argument("--log-level=3")
    option.add_argument('--disable-browser-side-navigation')
    option.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=option)
    # driver.implicitly_wait(60)

    wait = WebDriverWait(driver, 10)
    waitLong = WebDriverWait(driver, 30)
    waitLongLong = WebDriverWait(driver, 100)
    waitShort = WebDriverWait(driver, 5)
    waitShortShort = WebDriverWait(driver, 2)

    driver.delete_all_cookies()


if __name__ == "__main__":
    global fw
    global window
    global driver
    global wait
    global waitLong
    global waitLongLong
    global waitShort
    global waitShortShort
    global reload_interval_value

    if os.path.exists('log'):
        os.remove('log')

    logger = logging.getLogger(__file__)
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(funcName)s - %(lineno)d - %(levelname)s - %(message)s')
    fileHandler = logging.FileHandler('log')
    fileHandler.setFormatter(formatter)
    fileHandler.setLevel(logging.DEBUG)
    logger.addHandler(fileHandler)

    thread_flag = True

    max_refresh_attempt = 3

    init_driver()

    mainForm()



