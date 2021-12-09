__author__ = 'Fuzz'

from time import sleep
from RPA.Browser.Selenium import webdriver
from selenium.common.exceptions import *


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


def get_info_from_driver(driver, param, selector, attr='href'):
    """Receives 4 parameters, the web driver, the str to use for search (link text)
    and a selector for select strings or elements on the return.
    :returns a list with strings
    """
    sleep(3)
    elements_found = driver.find_elements_by_partial_link_text(str(param))
    output = []
    for element in elements_found:
        if selector == 's':
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
    investments_matrix = [[''] * cols for i in range(rows)]
    aux_col = aux_row = 1
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
            selector = driver.find_element_by_xpath("//select[@name='investments-table-object_length']")
            selector.send_keys('a')
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
