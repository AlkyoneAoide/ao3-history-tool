from faulthandler import is_enabled
import sys
from time import sleep
import xlsxwriter
import atexit
import signal
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

processedWorks = 0
warningAccepted = False

def help():
    print("Usage: 'python saveHistory.py <ao3 username> <ao3 password> <browser type> <y/n for including bookmark status>'")
    print("Acceptable browser types are: 'firefox'; 'chrome'")
    atexit.unregister(saveExit)
    sys.exit(2)

def saveExit():
    browser.get("https://archiveofourown.org/users/logout")
    browser.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div[2]/form/p[2]/input").click()

    browser.quit()
    book.close()

def exitSigHandler(a, b):
    sys.exit(0)

atexit.register(saveExit)
signal.signal(signal.SIGINT, exitSigHandler)

def logFromPage():

    works = ''
    try:
        works = browser.find_elements(By.XPATH, "/html/body/div[1]/div[2]/div[2]/ol[1]/li")
    except NoSuchElementException:
        sleep(1200)
        browser.refresh()
        works = browser.find_elements(By.XPATH, "/html/body/div[1]/div[2]/div[2]/ol[1]/li")

    window = browser.current_window_handle

    for work in works:
        global processedWorks
        global warningAccepted

        try:
            workURL = work.find_element(By.XPATH, ".//div[1]/h4/a[1]").get_attribute('href')
        except NoSuchElementException:
            sheet.write('A'+str((processedWorks+3)), work.find_element(By.XPATH, ".//div/h4").text)
            processedWorks += 1
            continue

        lastVisit = work.find_element(By.XPATH, ".//div[2]/h4").text.split()
        lastVisitConcat = lastVisit[2] + ' ' + lastVisit[3] + ' ' + lastVisit[4]
        rating = work.find_element(By.XPATH, ".//div[1]/ul/li[1]/a/span/span").text
        ratingAbbreviated = ''
        author = ''
        try:
            author = work.find_element(By.XPATH, ".//div[1]/h4/a[2]").text
        except NoSuchElementException:
            author = 'N/A'

        if rating.startswith("General"):
            ratingAbbreviated = 'G'
        elif rating.startswith("Teen"):
            ratingAbbreviated = 'T'
        elif rating.startswith("Mature"):
            ratingAbbreviated = 'M'
        elif rating.startswith("Explicit"):
            ratingAbbreviated = 'E'
        else:
            ratingAbbreviated = 'NR'

        sheet.write('A'+str((processedWorks+3)), lastVisitConcat)
        sheet.write('B'+str((processedWorks+3)), work.find_element(By.XPATH, ".//div[1]/h4/a[1]").text)
        sheet.write('C'+str((processedWorks+3)), author)
        sheet.write('D'+str((processedWorks+3)), work.find_element(By.XPATH, ".//dl/dd[2]").text)
        sheet.write('E'+str((processedWorks+3)), work.find_element(By.XPATH, ".//div[1]/h5").text)
        sheet.write('F'+str((processedWorks+3)), ratingAbbreviated)

        if bookmark:
            browser.switch_to.new_window('tab')
            browser.get(workURL)

            tmpWinHandle = browser.current_window_handle
            browser.switch_to.window(window)
            browser.switch_to.window(tmpWinHandle)

            if warningAccepted == False and not (ratingAbbreviated == 'G' or ratingAbbreviated == 'T'):
                browser.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/ul/li[1]/a").click()
                warningAccepted = True

            try:
                elem = WebDriverWait(browser, 30).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".expandable.secondary")))
            finally:
                browser.switch_to.default_content()
                if browser.find_element(By.CSS_SELECTOR, ".bookmark_form_placement_open").text.strip().lower() == "bookmark":
                    sheet.write('G'+str((processedWorks+3)), '✘')
                else:
                    sheet.write('G'+str((processedWorks+3)), '✔')

            browser.close()
            browser.switch_to.window(window)

            # Anti-DDOS protection
            # Change value (in seconds) to go slower or faster
            sleep(600)

        processedWorks += 1

if len(sys.argv) < 5:
    help()

ao3User = sys.argv[1] if type(sys.argv[1]) is str else help()
ao3Pass = sys.argv[2] if type(sys.argv[1]) is str else help()
browserType = sys.argv[3] if type(sys.argv[3]) is str else help()
bookmark = True
if sys.argv[4] == 'n':
    bookmark = False

book = xlsxwriter.Workbook('history.xlsx')
sheet = book.add_worksheet()

sheet.write('A1', 'Last Visited')
sheet.write('B1', 'Name')
sheet.write('C1', 'Author')
sheet.write('D1', 'Word Count')
sheet.write('E1', 'Fandom')
sheet.write('F1', 'Content Rating')
if bookmark:
    sheet.write('G1', 'Bookmarked?')

browser = webdriver.Firefox(executable_path=r'./geckodriver.exe')
if browserType.strip().lower() == 'chrome':
    browser = webdriver.Chrome(executable_path=r'./chromedriver.exe')

browser.get("https://archiveofourown.org/users/"+ ao3User +"/readings")

if browser.current_url == "https://archiveofourown.org/users/login":
    try:
        elem = WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.ID, "tos_agree")))
    finally:
        browser.find_element(By.ID, "tos_agree").click()
        sleep(1)
        browser.find_element(By.ID, "accept_tos").click()
        sleep(1)
        browser.find_element(By.ID, "user_login").send_keys(ao3User)
        browser.find_element(By.ID, "user_password").send_keys(ao3Pass)
        browser.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div[3]/div/form/dl/dd[4]/input").click()

browser.get("https://archiveofourown.org/users/"+ ao3User +"/readings")
maxPages = browser.find_elements(By.XPATH, "/html/body/div[1]/div[2]/div[2]/ol[2]/li")[-2].text

for i in range(1, int(maxPages)+1):
    browser.get("https://archiveofourown.org/users/"+ ao3User +"/readings?page="+ str(i))
    logFromPage()
    print(str(i)+"/"+maxPages)
    sleep(10)