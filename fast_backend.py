from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyQt5.QtCore import QObject, pyqtSignal
import pandas as pd
import sys
import os
import time

class Fast(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(int)
    error = pyqtSignal()
    result = pyqtSignal()

    def __init__(self, input_path, base_path, output_file_name):
        super(Fast, self).__init__()

        self.input_path = input_path
        self.base_path = base_path
        self.output_file_name = output_file_name

    def startup(self):
        self.driver_options = Options()
        self.driver_options.add_argument('--headless')

        self.fan_counter = 0

        # import username and password
        with open(self.resource_path('creds.txt'), 'r') as creds:
            self.username = creds.readline().strip('\n')
            self.password = creds.readline()

        #create empty lists for each attribute being exported into excel
        self.contact_names = []
        self.contact_types = []
        self.contact_days = []
        self.contact_emails = []
        self.contact_phones = []
    
    def resource_path(self, relative_path):
        return os.path.join(self.base_path, relative_path)

    def import_fans(self):
        # import data from excel
        business_and_fan_df = pd.read_excel(self.input_path, "Sheet1")
        self.business_names = business_and_fan_df['Business']
        self.fan_numbers = business_and_fan_df['Buid']        
    
    def create_driver(self):
        # create webdriver
        self.driver = webdriver.Edge(options=self.driver_options)
    
    def restart_driver(self):
        self.driver.close()
        self.driver = webdriver.Edge(options=self.driver_options)
    
    def login(self):
        self.driver.get('https://fansearch.web.att.com/fast/search?8&q')
        username = self.driver.find_element(By.ID, 'GloATTUID')
        password = self.driver.find_element(By.ID, 'GloPassword')
        log_on_button = self.driver.find_element(By.ID, 'GloPasswordSubmit')
        username.send_keys(self.username)
        password.send_keys(self.password)
        log_on_button.click()

        # FIXME change this to check for incorrect username or password specifically
        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'token-input')))
            self.result.emit()
        except:
            # if search bar is not found we are not logged in and must stop
            self.driver.close()
            self.error.emit()

        # Click continue button if password expiration notice occurs
        try:
            continue_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(By.ID, 'successButtonId'))
            continue_button.click()
        except Exception:
            # no password expiration notice
            pass

    def scrape_contacts(self):
        #main loop to grab contact data for each buid
        while self.fan_counter < len(self.business_names):
            try:
                print(self.fan_counter+1)
                self.driver.get('https://fansearch.web.att.com/fast/search?q')
                #find search bar and enter correct number
                search_bar = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'token-input')))
                search_bar.send_keys(str(self.fan_numbers[self.fan_counter]))

                #click search
                search_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/div[2]/div/form/div/div/div[2]/button')))
                search_button.click()

                # TODO keep track of the fans that are invalid
                # click account if the fan is invalid it will be skipped
                try:
                    account = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/div[3]/div/form/div[2]/table/tbody/tr/td[3]/div/a')))
                    account.click()
                except:
                    self.fan_counter += 1
                    self.progress.emit(self.fan_counter)
                    continue

                customer_contacts = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/span/div[1]/ul/li[4]/a')))
                customer_contacts.click()

                #add the business name to the columns for seperation
                self.contact_names.append(self.business_names[self.fan_counter])
                self.contact_types.append(self.fan_numbers[self.fan_counter])
                self.contact_days.append('')
                self.contact_emails.append('')
                self.contact_phones.append('')

                #loop to add name, type, email, and phone number into seperate list to be put in a data frame
                for name in WebDriverWait(self.driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div[2]/div/div/span/div[2]/form[1]/table/tbody/tr/td[1]/div'))):
                    self.contact_names.append(name.text)
                for type in WebDriverWait(self.driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div[2]/div/div/span/div[2]/form[1]/table/tbody/tr/td[2]/div'))):
                    self.contact_types.append(type.text)
                for day in WebDriverWait(self.driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div[2]/div/div/span/div[2]/form[1]/table/tbody/tr/td[6]/div'))):
                    self.contact_days.append(day.text)
                for email in WebDriverWait(self.driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div[2]/div/div/span/div[2]/form[1]/table/tbody/tr/td[3]/div'))):
                    self.contact_emails.append(email.text)
                for phone in WebDriverWait(self.driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div[2]/div/div/span/div[2]/form[1]/table/tbody/tr/td[10]/div'))):
                    self.contact_phones.append(phone.text)
                home_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/ul[1]/li[1]/a')))
                home_button.click()

                self.fan_counter += 1

            except Exception as e:
                print(e)
                self.restart_driver()
                self.login()

            self.progress.emit(self.fan_counter)
        self.finished.emit()

    def export_contacts_to_excel(self):
        #create pandas dataframe of Name type email and phone
        df = pd.DataFrame({"Name": self.contact_names, "Type": self.contact_types, 'Days': self.contact_days, 'Email': self.contact_emails, 'Phone': self.contact_phones})
        #write dataframe created to excel sheet
        df.to_excel(self.output_file_name, sheet_name='Fast Contacts')
        print('done')

if __name__ == "__main__":
    fast = Fast()
    fast.base_path = ''
    fast.output_file_name = 'Fast Contacts.xlsx'
    fast.input_path = "C:/Users/av037e/OneDrive - AT&T Services, Inc/Desktop/coding/fast_input.xlsx"
    fast.startup()
    fast.import_fans()
    fast.create_driver()
    fast.login()
    fast.scrape_contacts()
    fast.export_contacts_to_excel()