__author__ = 'Fuzz'
from selenium import webdriver
from selenium.common import exceptions


def start_driver(browser: str, download_path=None, **options):
    """Start a webdriver with the given options."""
    browser = browser.strip()
    factory = getattr(webdriver, browser, None)

    if not factory:
        raise ValueError(f"Unsupported browser: {browser}")

    if download_path is None:
        pass
        driver = factory(**options)
    else:
        if browser == 'chrome':
            try:
                option = webdriver.ChromeOptions()
                option_info = {
                    "download.default_directory": download_path.strip()
                }
                option.add_experimental_option("prefs", option_info)
                driver = factory(chrome_options=option)
            except exceptions.WebDriverException:
                driver = factory(**options)
                print('problem setting path, local path set as default')
    return driver
