from selenium import webdriver
import os.path
import csv
import time
from selenium.common.exceptions import NoSuchElementException
import copy

def find_data(company_name): #function for downloading a file with company data (company_name)

    driver.implicitly_wait(10)
    driver.get('https://finance.yahoo.com/')

    search_field = driver.find_element_by_xpath('//*[@id="yfin-usr-qry"]')
    search_field.send_keys(company_name)

    time.sleep(1)
    try:
        check_presence_company = driver.find_element_by_xpath('//*[@id="header-search-form"]/div[2]/div[1]/div/div[1]/h3')
    except NoSuchElementException:
        company = 'exception' # if the page or element did not load
        return company

    if check_presence_company.text == 'Symbols':
        try:
            search_button = driver.find_element_by_xpath('//*[@id="header-search-form"]/div[2]/div[1]/div/ul[1]/li[1]')
            search_button.click()

            historical_data = driver.find_element_by_xpath('//*[@data-test="HISTORICAL_DATA"]/a')
            historical_data.click()

            time_period = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/div[1]/div/div')
            time_period.click()

            button_Max = driver.find_element_by_xpath('//*[@data-value="MAX"]')
            button_Max.click()

            button_apply = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[1]/button')
            button_apply.click()

            download_file = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[1]/div[2]/span[2]/a')
            download_file.click()
        except NoSuchElementException:
            company = "exception" # if the page or element did not load
            return company
        company = "found" # Company is found and uploaded file
        return company
    else:
        company = "not_found" # Company is not found
        return company

def check_files_with_data(list_of_companies, not_found = []): # function to check if data files exist

    for company_name in list_of_companies: # check all companies
        file_name = company_name + '.csv'
        path_to_file = 'C:\\Users\Public\Downloads\\' + file_name
        time.sleep(1)

        if os.path.exists(path_to_file): # if file uploaded -> continue
            continue
        else: # if file is not loaded -> download
            find = find_data(company_name)

            if find == 'not_found': # if company is not found -> remove this company from list
                not_found.append(company_name)
                list_of_companies.remove(company_name)
            elif find == 'exception': # if have exception -> skip this company
                list_of_companies.remove(company_name)
            check_files_with_data(list_of_companies, not_found)
    return not_found # return the list of companies which not be found

def calculate_change(list_of_companies):

    for company_name in list_of_companies:
        file_name = company_name + '.csv'
        path_to_file = 'C:\\Users\Public\Downloads\\' + file_name

        if os.path.exists(path_to_file):  # if data file exist -> read the data
            with open(path_to_file, 'r') as file:
                table = csv.reader(file)
                table_of_data = list(table)
            titles = table_of_data[0]
            titles.append('3day_before_change') # create titles
            table_of_data = table_of_data[1:] # create table with data
            table_of_data.reverse() # start from last date

            # convert the date to integer form
            for row in table_of_data: #
                row[0] = row[0].split('-')

                for i in range(len(row[0])):
                    row[0][i] = int(row[0][i])

            # creating a new column
            for row in table_of_data:
                year = row[0][0]
                month = row[0][1]
                day = row[0][2]
                days_in_month = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31] # helper

                # find out the date three days ago
                if year % 4 == 0:
                    days_in_month[1] = 29

                if day < 4 and month < 2:
                    year -= 1
                    month = 12
                    day = 31 + day - 3
                elif day < 4:
                    month -= 1
                    day = days_in_month[month - 1] + day - 3
                else:
                    day = day - 3

                # check this date in the data
                for row_check in table_of_data:
                    if [year,month,day] in row_check: # if date exist -> calculate ratio
                        row.append(float(row[4])/float(row_check[4]))

                if len(row) == 7: # if date not exist -> -
                    row.append('-')

            # convert the date to string form
            for row in table_of_data:
                row[0][0] = str(row[0][0])
                if row[0][1] < 10:
                    row[0][1] = str("0") + str(row[0][1])
                else:
                    row[0][1] = str(row[0][1])

                if row[0][2] < 10:
                    row[0][2] = str("0") + str(row[0][2])
                else:
                    row[0][2] = str(row[0][2])

                row[0] = row[0][0] + str('-') + row[0][1] + str('-') + row[0][2]

            table_of_data.reverse() # adding a titles
            table_of_data.append(titles)
            table_of_data.reverse()

            with open(file_name, "w", newline="\n") as file: # saving in new file
                writer = csv.writer(file)
                writer.writerows(table_of_data)

def find_summary(company_name):
    driver.implicitly_wait(10)
    table_with_summary_info = []
    table_with_summary_info.append(['link', 'title '])

    driver.get('https://finance.yahoo.com/')

    search_field = driver.find_element_by_xpath('//*[@id="yfin-usr-qry"]')
    search_field.send_keys(company_name)

    time.sleep(1)
    try:
        check_presence_company = driver.find_element_by_xpath('//*[@id="header-search-form"]/div[2]/div[1]/div/div[1]/h3')
    except NoSuchElementException:
        company = "exception" # if the page or element did not load
        return company

    if check_presence_company.text == "Symbols":
        try:
            search_button = driver.find_element_by_xpath('//*[@id="header-search-form"]/div[2]/div[1]/div/ul[1]/li[1]')
            search_button.click()

            summary = driver.find_element_by_xpath('//*[@data-test="SUMMARY"]/a')
            summary.click()

            news = driver.find_element_by_xpath('//*[@id="Col1-3-Summary-Proxy"]/section/div/div/a[1]')
            news.click()

            latest_news = driver.find_element_by_xpath('//h3/a[@target="_self"]')
            link = latest_news.get_attribute('href')
            title = latest_news.text
            table_with_summary_info.append([link, title])

            file_name = company_name + '_summary.csv'
        except NoSuchElementException:
            company = 'exception' # if the page or element did not load
            return company

        with open(file_name, "w", newline="\n") as file:
            writer = csv.writer(file)
            writer.writerows(table_with_summary_info)
        company = 'found' # Company is found and uploaded file
        return company
    else:
        company = 'not_found' # Company is not found
        return company

def check_files_with_summary(list_of_companies, not_found = []):

    for company_name in list_of_companies:
        file_name = company_name + '_summary.csv'
        time.sleep(1)

        if os.path.exists(file_name):
            continue
        else:
            find = find_summary(company_name)

            if find == 'not_found': # if company is not found -> remove this company from list
                not_found.append(company_name)
                list_of_companies.remove(company_name)
            elif find == 'exception': # if have exception -> skip this company
                list_of_companies.remove(company_name)
            check_files_with_summary(list_of_companies, not_found)
    return not_found # return the list of companies which not be found


list_of_companies = ['PD', 'ZUO', 'PINS', 'ZM', 'PVTL', 'DOCU', 'CLDR', 'RUN']
not_found = []

if __name__ == '__main__':

    while len(list_of_companies) > 0:
        chromeOptions = webdriver.ChromeOptions()
        prefs = {"download.default_directory" : r'C:\Users\Public\Downloads'} # specify a new download directory (available on Windows)
        chromeOptions.add_experimental_option('prefs',prefs)

        driver = webdriver.Chrome(chrome_options=chromeOptions)
        driver.implicitly_wait(10)

        list_ = copy.deepcopy(list_of_companies)

        not_found = check_files_with_data(list_, not_found)

        for comp in not_found:  # if company is not found -> remove this company from list
            if comp in list_of_companies:
                list_of_companies.remove(comp)

        calculate_change(list_of_companies)

        list_ = copy.deepcopy(list_of_companies)

        not_found = check_files_with_summary(list_, not_found)

        for comp in not_found: # if company is not found -> remove this company from list
            if comp in list_of_companies:
                list_of_companies.remove(comp)

        for name in list_of_companies: # remove a company from list if all data about it is saved
            if os.path.exists(name + '.csv') and os.path.exists(name + '_summary.csv'):
                list_of_companies.remove(name)

        driver.close()



