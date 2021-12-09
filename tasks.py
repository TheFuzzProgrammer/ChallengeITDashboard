__author__ = 'Fuzz'

import sys
from time import sleep
from RPA.Browser.Selenium import webdriver
from selenium.common.exceptions import *
from RPA.Excel.Files import *


def main(args=None):
    """
    RPA CHALLENGE main MODULE
    uses:
    >>$ python main.py 'agency that need'
    END CODES:
    0 all was fine
    1 you don't send parameters and decided to exit
    2 It couldn't found the agency
    """
    # Get Web driver space
    driver = get_driver('Chrome', 'https://itdashboard.gov/')
    sleep(5)
    # Get first button
    dive_in_button = driver.find_element_by_link_text('DIVE IN')
    dive_in_button.click()
    # Get view buttons
    view_buttons = driver.find_elements_by_link_text('view')
    view_data = get_info_from_driver(driver, '$', 's', 'text')
    links = get_urls(view_buttons)

    if args is None:
        matrix = get_all_from_driver(links, driver)
        path = 'output\investment' + str(matrix[0][0][0]) + '.xlsx'
        all_to_excel_file(matrix, path)
    else:
        matrix = get_from_agency_page(driver, args)
        if matrix == 0:
            print('Agency not found')
            exit(2)
        else:
            path = 'output\investment' + str(matrix[0][0]) + '.xlsx'
            save_to_excel(matrix, path)
    info = Elements(links, view_data)
    info.investments_matrix = matrix
    info.show_data_as_string()
    info.show_link_as_string()
    driver.close()


def all_to_excel_file(matrix, path=None):
    for table in range(0, len(matrix)):
        save_to_excel(matrix[table], path)
    return 0


class Elements(object):
    """ contains elements from driver, an instance need:
        a list of links, and a list of data, use it as you need
    """

    def __init__(self, links=None, data=None):
        super().__init__()
        if links is None:
            links = []
        if data is None:
            data = []
        self._is_link = links
        self._is_data = data
        self.investments_matrix = []
        self.business_cases_urls = []

    def show_link_as_string(self):
        for element in self._is_link:
            if type(element) == str:
                print(element)
            else:
                print(str(element.get_attribute('text')) + '\n')
        return 0

    def show_data_as_string(self):
        for element in self._is_data:
            if type(element) == str:
                print(element)
            else:
                print(str(element.get_attribute('text')) + '\n')
        return 0


def get_info_from_driver(driver, param, selector_var, attr='href'):
    """Receives 4 parameters, the web driver, the str to use for search (link text)
    and a selector for select strings or elements on the return.
    :returns a list with strings
    """
    sleep(3)
    elements_found = driver.find_elements_by_partial_link_text(str(param))
    output = []
    for element in elements_found:
        if selector_var == 's':
            try:
                output.append(str(element.get_attribute(str(attr))))
            except AttributeError:
                print('There is not elements with ', str(attr), ' as attribute')
            finally:
                if element == elements_found[-1]:
                    return output
        else:
            output.append(element)
            if element == elements_found[-1]:
                return output


def get_all_from_driver(links, driver):
    tables = []
    for link in range(0, len(links)):
        print(link, ' OF ', len(links))
        driver.get(links[link])
        selector(driver)
        rows = len(driver.find_elements_by_xpath('//*[@id="investments-table-object"]/tbody/tr'))
        while rows == 10:
            rows = len(driver.find_elements_by_xpath('//*[@id="investments-table-object"]/tbody/tr'))
        cols = len(driver.find_elements_by_xpath('//*[@id="investments-table-object"]/tbody/tr[1]/td'))
        table = get_investments_table(driver, rows, cols)
        get_pdf(driver, get_investment_id(table[0][0]))
        tables.append(table)

    return list(tables)


def get_urls(web_elements_array):
    urls_array = []
    for x in web_elements_array:
        urls_array.append(x.get_attribute('href'))
    urls_array = set(urls_array)
    return list(urls_array)


def get_driver(browser, url):
    web_driver = webdriver.start(str(browser))
    web_driver.get(str(url))
    return web_driver


def get_investments_table(driver, rows, cols):
    investments_matrix = [[''] * cols for _ in range(rows)]
    elements = driver.find_elements_by_xpath('//*[@id="investments-table-object"]/tbody/tr')
    for element in range(0, len(elements)):
        investments_matrix[element][0] = elements[element].find_elements_by_tag_name('td')[0].text
        investments_matrix[element][1] = elements[element].find_elements_by_tag_name('td')[1].text
        investments_matrix[element][2] = elements[element].find_elements_by_tag_name('td')[2].text
        investments_matrix[element][3] = elements[element].find_elements_by_tag_name('td')[3].text
        investments_matrix[element][4] = elements[element].find_elements_by_tag_name('td')[4].text
        investments_matrix[element][5] = elements[element].find_elements_by_tag_name('td')[5].text
        investments_matrix[element][6] = elements[element].find_elements_by_tag_name('td')[6].text
        print('ROW : ', element + 1)

    return investments_matrix


def get_investment_id(string):
    string_id = ''
    for x in range(len(string)):
        if string[x] == '-':
            return string_id
        else:
            string_id += string[x]


def get_pdf(driver, investment_id):
    urls = get_urls(driver.find_elements_by_partial_link_text(investment_id))
    for x in urls:
        driver.get(x)
        button_found = False
        while not button_found:
            try:
                pdf = driver.find_element_by_partial_link_text('PDF')
                pdf.click()
                button_found = True
                sleep(5)
            except NoSuchElementException:
                pass
    return 0


def selector(driver):
    selector_found = False
    while not selector_found:
        try:
            driver.find_element_by_xpath("//select[@name='investments-table-object_length']").send_keys('a')
            selector_found = True
        except:
            continue


def get_from_agency_page(driver, string):
    axis_y = len(driver.find_elements_by_xpath('//*[@id="agency-tiles-widget"]/div/div'))
    axys_x = len(driver.find_elements_by_xpath('//*[@id="agency-tiles-widget"]/div/div[1]/div'))

    for row in range(1, axis_y+1):
        for column in range(1, axys_x+1):
            agency = driver.find_element_by_xpath(
                '//*[@id="agency-tiles-widget"]/div/div['+str(row)+']/div['+str(column)+']')
            if string.upper() in agency.text.upper():
                agency.click()
                selector(driver)
                rows = len(driver.find_elements_by_xpath('//*[@id="investments-table-object"]/tbody/tr'))
                while rows == 10:
                    rows = len(driver.find_elements_by_xpath('//*[@id="investments-table-object"]/tbody/tr'))
                cols = len(driver.find_elements_by_xpath('//*[@id="investments-table-object"]/tbody/tr[1]/td'))
                table = get_investments_table(driver, rows, cols)
                get_pdf(driver, get_investment_id(table[0][0]))
                return table
            else:
                pass
        return 0


def read_excel_worksheet(path, worksheet):
    lib = Files()
    lib.open_workbook(path)
    lib.create_workbook(path, "xlsx")
    try:
        return lib.read_worksheet(worksheet)
    finally:
        lib.close_workbook()


def save_to_excel(table, path=None, exists=False):
    """ARGS
        table: A list() object format table[rows][cols]
        path: your file path, it will contain file name and extension
        exist: while False, it will take generic filename name TAKE CARE!
        """
    workbook = Files()
    if path is None:
        path = 'file.xlsx'
    if exists:
        workbook.open_workbook(path)
    else:
        workbook.create_workbook(path)

    workbook.append_rows_to_worksheet(table)
    workbook.save_workbook()
    workbook.close_workbook()

    return 0


if __name__ == '__main__':
    main('Agency')
