__author__ = 'Fuzz'

import sys
from elements import *
from file_handler import *


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
            save_to_excel(matrix)
    info = Elements(links, view_data)
    info.investments_matrix = matrix
    info.show_data_as_string()
    info.show_link_as_string()

    return info


def all_to_excel_file(matrix, path=None):
    for table in range(0, len(matrix)):
        save_to_excel(matrix[table], path)
    return 0


if __name__ == '__main__':
    print(sys.argv)
    if len(sys.argv) != 2:
        print('No argument setted, or its invalid\n '
              'WORKING WITH ARGUMENTS:\n<this script> "argument of execution" on terminal'
              'Without argument, it will take data from ALL the sources, continue? Y/N')
        option = input().upper()
        if option == 'Y':
            main()
        else:
            exit(1)
    else:
        main(sys.argv[1])
