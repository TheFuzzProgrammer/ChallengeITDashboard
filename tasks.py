__author__ = 'Fuzz'

from time import sleep
from SeleniumLibrary.errors import *
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import *
from RPA.PDF import PDF


def main():
    """
    RPA CHALLENGE main MODULE
    """
    # Get Web driver space
    driver = get_driver('https://itdashboard.gov/')
    sleep(3)
    driver.click_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a')
    agency_name = get_agency()
    table = get_from_agency_page(driver, agency_name)
    save_to_excel(table, 'output\info_from_'+agency_name+'.xlsx')


def get_driver(url):
    web_driver = Selenium()
    web_driver.open_available_browser(url)
    return web_driver


def get_investments_table(driver, rows, cols):
    investments_matrix = [[''] * cols for _ in range(rows)]
    elements = driver.get_webelements('//*[@id="investments-table-object"]/tbody/tr')
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


def get_pdf(driver, rows):
    urls = get_urls(driver, rows)
    for x in urls:
        driver.go_to(x)
        button_found = False
        while not button_found:
            try:
                pdf = driver.get_webelement('//*[@id="business-case-pdf"]/a')
                pdf.click()
                button_found = True
                sleep(5)
            except ElementNotFound:
                pass
    return 0


def selector(driver):
    selector_found = False
    while not selector_found:
        try:
            driver.get_webelement("//select[@name='investments-table-object_length']").send_keys('a')
            selector_found = True
        except ElementNotFound:
            pass


def get_urls(driver, rows):
    urls_array = []
    for x in range(1, rows+1):
        try:
            web_element = driver.get_webelement('//*[@id="investments-table-object"]/tbody/tr[' + str(x) + ']/td/a')
            urls_array.append(web_element.get_attribute('href'))
        except ElementNotFound:
            pass

    return urls_array


def get_from_agency_page(driver, agency):
    axis_y = len(driver.get_webelements('//*[@id="agency-tiles-widget"]/div/div'))
    axis_x = len(driver.get_webelements('//*[@id="agency-tiles-widget"]/div/div[1]/div'))
    for row in range(1, axis_y + 1):
        for column in range(1, axis_x + 1):
            agency_found = driver.get_webelement(
                '//*[@id="agency-tiles-widget"]/div/div[' + str(row) + ']/div[' + str(column) + ']')
            if agency.upper() in agency_found.text.upper():
                agency_found.click()
                selector(driver)
                rows = len(driver.get_webelements('//*[@id="investments-table-object"]/tbody/tr'))
                while rows == 10:
                    rows = len(driver.get_webelements('//*[@id="investments-table-object"]/tbody/tr'))
                cols = len(driver.get_webelements('//*[@id="investments-table-object"]/tbody/tr[1]/td'))
                table = get_investments_table(driver, rows, cols)
                get_pdf(driver, rows)
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


def get_agency():
    data = open('source.txt', 'r')
    agency = data.read()
    data.close()
    return str(agency)


def compare_pdf(path, investment_id):
    pdf = PDF()
    text = str(pdf.get_text_from_pdf(path, '1'))
    if investment_id in text:
        return True
    else:
        return False


if __name__ == '__main__':
    main()
