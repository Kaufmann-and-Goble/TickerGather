from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

import time
import datetime
import os
import sys

cwd = '/Volumes/Shared/Apps/morningstar'
ticker = []
errors = []
curtime = datetime.datetime.now()
filename = curtime.strftime("%Y-%m-%d")
imagedir = (cwd + '/' + filename + '/images/')
textdir = (cwd + '/' + filename + '/text/')
pdfdir = (cwd + '/' + filename + '/pdfs/')
printname = 'HP_LaserJet_600_M602__Ken_'

# Get the list of ticker symbols


def setup():
    file = open(cwd + '/tickers')
    dupe = 0
    for line in file:
        line = str(line).replace('\n', '')
        if line not in ticker:
            ticker.append(line)
        else:
            dupe += 1
    print('Dupes = ' + str(dupe))

# Open Chrome, navigate to website, click Month tab, Grab table data, Resize Windows, Snap Screenshot
# For ticker list


def getdata(count):
    holder = []
    table = []
    count1 = 0
    count3 = 0
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    browser = webdriver.Chrome(cwd + '/chromedriver', chrome_options=chrome_options)
    browser.set_window_size(900, 1230)
    page = "http://performance.morningstar.com/fund/performance-return.action?ops=p&p=total_returns_page&t=" + str(ticker[count]) + "&region=usa&culture=en-US&s=0P0000J533"
    browser.get(page)

    try:
        browser.find_element_by_xpath('/html/body/div/div/div[2]/ul[1]/li[2]/a').click()
    except NoSuchElementException:
        print('Error with ' + str(ticker[count]))
        errors.append(str(ticker[count]))
        count += 1

    try:
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="tab-month-end-content"]/table')))
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="plan529TotalReturn"]/div/canvas[2]')))
    except:
        print('Could not find either table or graph for ' + str(ticker[count]))

    elements = browser.find_elements_by_tag_name('tr')

    for ele in elements:
        holder.append(ele.get_attribute('innerHTML'))

    while count1 < len(holder):

        if 'Total Return' in str(holder[count1]) and '6-Month' in str(holder[count1]):
            table.append(holder[count1])
            count2 = count1 + 1

            while 'row_data' in holder[count2]:
                table.append(holder[count2])
                count2 += 1
                count1 += 1

        count1 += 1

# Clean up the HTML taken from the website

    while count3 < len(table):
        table[count3] = str(table[count3].replace('\n', ''))
        table[count3] = str(table[count3].replace('\t', ''))
        table[count3] = str(table[count3].replace('<th scope="row" class="row_lbl" style="word-break:break-all;overflow:hidden; ">', ''))
        table[count3] = str(table[count3].replace('<td class="row_data">', ''))
        table[count3] = str(table[count3].replace('</td>', '\t'))
        table[count3] = str(table[count3].replace('</th>', '\t'))
        table[count3] = str(table[count3].replace('</span>', ''))
        table[count3] = str(table[count3].replace('<span>', ''))
        table[count3] = str(table[count3].replace('<th scope="row" class="row_lbl divide" style="word-break:break-all;overflow:hidden; ">', ''))
        table[count3] = str(table[count3].replace('<td class="row_data_0">', ''))
        table[count3] = str(table[count3].replace('<th scope="row" class="col_head_lbl">', ''))
        table[count3] = str(table[count3].replace('<th scope="col" class="col_data">', ''))
        table[count3] = str(table[count3].replace('<th scope="col" class="col_data_0">', ''))
        table[count3] = str(table[count3].replace('<td class="row_data divide">', ''))
        table[count3] = str(table[count3].replace('<td class="row_data_0 divide">', ''))
        table[count3] = str(table[count3].replace('<th scope="row" class="row_lbl divide">', ''))
        table[count3] = str(table[count3].replace('&nbsp', ''))
        count3 += 1

# Create text save dirs

    if not os.path.exists(textdir):
        os.makedirs(textdir)
        print('Creating Directory for text.')

    file = open(textdir + str(ticker[count]) + '.txt', 'w+')

# Resize

    for line in table:
        file.write(line + '\n')

    file.close()
    browser.execute_script("document.body.style.zoom='90%'")

# Create image save dirs

    if not os.path.exists(imagedir):
        os.makedirs(imagedir)
        print('Creating Directory for images.')

    browser.get_screenshot_as_file(imagedir + str(ticker[count]) + '.png')
    browser.close()

# Run the main program


def run():
    total_start_time = time.time()
    count = 0
    tally = 1
    while count < len(ticker):
        start_time = time.time()
        getdata(count)
        finish_time = time.time() - start_time
        print(str(ticker[count]) + ' [DONE] in ' + str(finish_time) + ' seconds.')
        print('[Progress] = ' + str(round((tally/len(ticker)) * 100)) + '%')
        sys.stdout.write("\033[F")
        sys.stdout.write("\033[K")

        if tally == len(ticker):
            print('[Finished] [Cleaning Up] 100%')

        count += 1
        tally += 1
    total_finish_time = time.time() - total_start_time
    print('Total time = ' + str(total_finish_time) + ' seconds')
    prompt()

# Open Chrome, navigate to website, click Month tab, Grab table data, Resize Windows, Snap Screenshot
# For single ticker


def singleRun(request):
    start_time = time.time()
    holder = []
    table = []
    count1 = 0
    count3 = 0
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    browser = webdriver.Chrome(cwd + '/chromedriver', chrome_options=chrome_options)
    browser.set_window_size(900, 1230)
    page = "http://performance.morningstar.com/fund/performance-return.action?ops=p&p=total_returns_page&t=" + str(
        request) + "&region=usa&culture=en-US&s=0P0000J533"
    browser.get(page)

    try:
        browser.find_element_by_xpath('/html/body/div/div/div[2]/ul[1]/li[2]/a').click()
    except NoSuchElementException:
        print('Error with ' + request)
        errors.append(request)

    try:
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="tab-month-end-content"]/table')))
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="plan529TotalReturn"]/div/canvas[2]')))
    except:
        print('Could not find either table or graph for ' + request)

    elements = browser.find_elements_by_tag_name('tr')

    for ele in elements:
        holder.append(ele.get_attribute('innerHTML'))

    while count1 < len(holder):

        if 'Total Return' in str(holder[count1]) and '6-Month' in str(holder[count1]):
            table.append(holder[count1])
            count2 = count1 + 1

            while 'row_data' in holder[count2]:
                table.append(holder[count2])
                count2 += 1
                count1 += 1

        count1 += 1

# Clean up the HTML taken from the website

    while count3 < len(table):
        table[count3] = str(table[count3].replace('\n', ''))
        table[count3] = str(table[count3].replace('\t', ''))
        table[count3] = str(table[count3].replace('<th scope="row" class="row_lbl" style="word-break:break-all;overflow:hidden; ">',''))
        table[count3] = str(table[count3].replace('<td class="row_data">', ''))
        table[count3] = str(table[count3].replace('</td>', '\t'))
        table[count3] = str(table[count3].replace('</th>', '\t'))
        table[count3] = str(table[count3].replace('</span>', ''))
        table[count3] = str(table[count3].replace('<span>', ''))
        table[count3] = str(table[count3].replace('<th scope="row" class="row_lbl divide" style="word-break:break-all;overflow:hidden; ">', ''))
        table[count3] = str(table[count3].replace('<td class="row_data_0">', ''))
        table[count3] = str(table[count3].replace('<th scope="row" class="col_head_lbl">', ''))
        table[count3] = str(table[count3].replace('<th scope="col" class="col_data">', ''))
        table[count3] = str(table[count3].replace('<th scope="col" class="col_data_0">', ''))
        table[count3] = str(table[count3].replace('<td class="row_data divide">', ''))
        table[count3] = str(table[count3].replace('<td class="row_data_0 divide">', ''))
        table[count3] = str(table[count3].replace('<th scope="row" class="row_lbl divide">', ''))
        table[count3] = str(table[count3].replace('&nbsp', ''))
        count3 += 1

# Create text save dirs

    if not os.path.exists(textdir):
        os.makedirs(textdir)
        print('Creating Directory for text.')

    file = open(textdir + str(request) + '.txt', 'w+')

# Resize

    for line in table:
        file.write(line + '\n')

    file.close()
    browser.execute_script("document.body.style.zoom='90%'")

# Create image save dirs

    if not os.path.exists(imagedir):
        os.makedirs(imagedir)
        print('Creating Directory for images.')

    browser.get_screenshot_as_file(imagedir + str(request) + '.png')
    browser.close()
    elapsed_time = time.time() - start_time
    print('Finished ' + request + ' in ' + str(elapsed_time) + ' seconds.')

# Prompt the user for options.


def convert():
    if not os.path.exists(pdfdir):
        os.makedirs(pdfdir)
        print('Creating Directory for pdfs.')
    for item in ticker:
        print('Converting ' + item)
        os.system('convert ' + imagedir + item + '.png ' + pdfdir + item + '.pdf')
    print('Conversions [DONE]')
    prompt()


def prompt():
    text = '----------------------------------------\n--Return RUN to start data gather.'
    text = text + '\n--Return RERUN to run a single ticker again.'
    text = text + '\n--Return CONVERT to convert all scanned images to PDFs.'
    text = text + '\n--Return INFO to see the list of tickers to be gathered.'
    text = text + '\n--Return CHANGE to change the ticker list.'
    text = text + '\n--Return SHOW to view gathered data folders.'
    text = text + '\n--Return PRINT to print all scanned images.'
    text = text + '\n--Return QUIT to quit program.\n----------------------------------------\n= '
    start = input(text)

    if start.lower() == 'run':
        print('Starting...\n')
        run()

    if start.lower() == 'rerun':
        single = input('Please type the ticker here: ')
        singleRun(single)
        prompt()

    if start.lower() == 'info':
        del ticker[:]
        setup()
        for item in ticker:
            print(str(item))
        print('Total = ' + str(len(ticker)))
        prompt()

    if start.lower() == 'change':
        os.system('open ' + cwd + '/tickers')
        del ticker[:]
        setup()
        prompt()

    if start.lower() == 'show':
        os.system('open ' + imagedir)
        os.system('open ' + textdir)
        prompt()

    if start.lower() == 'print':
        for line in ticker:
            os.system('lpr -P ' + printname + ' ' + imagedir + line + '.png')
            print('Printing ' + imagedir + line + '.png to ' + printname)
        prompt()

    if start.lower() == 'convert':
        convert()

    if start.lower() == 'quit':
        print('Goodbye.')
        quit()

setup()
prompt()


# Andrew Zimdars (Kaufmann and Goble Associates)
