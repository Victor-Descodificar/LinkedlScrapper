"""
Created on Aug 14, 2018
@author: Victor Ferreira Pereira
@contact: vfpereira@gmail.com
"""

import json
import os
import time
from random import randint

import xlsxwriter
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class LinkedlnScraper:
    DOWN_CLICK = 3

    def __init__(self):
        # -- Read JSON file
        with open(os.path.join(os.path.dirname(__file__), "config.json"), 'r') as json_data:
            self.credentials = json.load(json_data)

        # -- Start opening the browser
        first_link = 'https://www.linkedin.com/sales/search'

        _profile = webdriver.FirefoxProfile()
        _profile.set_preference("browser.cache.disk.enable", False)
        _profile.set_preference("browser.cache.memory.enable", False)
        _profile.set_preference("browser.cache.offline.enable", False)
        _profile.set_preference("network.http.use-cache", False)
        self.driver = webdriver.Firefox(_profile)
        self.driver.get(first_link)

        # Spreadsheet headers
        self.user_name = ['Name of Contact']
        self.url_name = ['LinkedIn profile URL']
        self.current_title = ['Current title']
        self.company_name = ['Current Company']
        self.url_company = ['Company\'s LinkedIn URL']
        self.not_found_message = 'No Information Available'
        self.end_loop = False

    def wait_element(self, e):
        try:
            element_present = EC.presence_of_element_located((By.ID, e))
            WebDriverWait(self.driver, 30).until(element_present)
        except TimeoutException:
            raise Exception("Timed out loading page")

    def move_page_down(self):
        for i in range(LinkedlnScraper.DOWN_CLICK):
            self.driver.execute_script('window.scrollTo(0,100000);')
            time.sleep(1)

    def move_next(self):
        self.move_page_down()
        if [] != self.driver.find_elements_by_css_selector('a.next-pagination.page-link.disabled'):
            return True

        time.sleep(randint(5, 10))
        self.driver.find_element_by_css_selector('a.next-pagination.page-link').click()

    def login(self):
        self.wait_element('session_key-login')
        self.driver.find_element_by_id('session_key-login').send_keys(self.credentials['user'])
        self.driver.find_element_by_id(
            'session_password-login').send_keys(self.credentials['passwd'])
        self.driver.find_element_by_id('btn-primary').click()

        # -- Open the search view
        self.driver.find_element_by_partial_link_text('View all filters').click()

        # -- Wait the user type the search
        while True:
            if self.driver.find_elements_by_class_name('modal-header'):
                time.sleep(2)
            else:
                break

    def populate_info(self):

        while True:
            self.move_page_down()
            # Count how many profiles there are in the search
            info = self.driver.find_elements_by_class_name('content-wrapper')
            # Get size
            size = len(info)

            for i in range(0, size):
                # Get the name and URL from profile
                name = info[i].find_elements_by_class_name('name-link')
                if name != []:
                    self.user_name.append(name[0].text)
                    self.url_name.append(str(name[0].get_attribute('href')))
                else:
                    self.user_name.append(self.not_found_message)
                    self.url_name.append(self.not_found_message)

                # Get company name and URL from profile
                company = info[i].find_elements_by_class_name('company-name')
                if company != []:
                    self.company_name.append(company[0].text)
                    self.url_company.append(str(company[0].get_attribute('href')))
                else:
                    self.company_name.append(self.not_found_message)
                    self.url_company.append(self.not_found_message)

                # Get current title from profile
                title = info[i].find_elements_by_class_name('info-value')
                if title != []:
                    self.current_title.append(title[0].text)
                else:
                    self.current_title.append(self.not_found_message)

            # Next page
            end_loop = self.move_next()
            if end_loop:
                break

    # --- Create a XLSX file with names
    def write_spreadsheet(self):
        final_list = [self.company_name, self.url_company, self.user_name, self.current_title, self.url_name]
        workbook = xlsxwriter.Workbook('Names.xlsx')
        worksheet = workbook.add_worksheet()
        row = 0
        for col, data in enumerate(final_list):
            worksheet.write_column(row, col, data)

        workbook.close()

    def close(self):
        print("Data collection completed.")
        self.driver.quit()


if __name__ == '__main__':
    scrapper = LinkedlnScraper()
    scrapper.login()
    scrapper.populate_info()
    scrapper.write_spreadsheet()
    scrapper.close()
