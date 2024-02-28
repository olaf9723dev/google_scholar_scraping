import os
import requests
import csv
import random
from openpyxl import load_workbook
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

MAIN_URL = "https://scholar.google.com/"
DATA_PATH = "config/testing_data.xlsx"
PROXIES = 'config/proxy_ips.xlsx'

class Extractor:
    def __init__(self) -> None:
        self.keys = []
        self.links = []
        self.session = requests.Session()
        self.proxies = []
        try:
            os.mkdir('output')
        except FileExistsError:
            pass

    def get_data(self):
        self.read_keys(DATA_PATH)
        self.get_links()

    def get_links(self):
        query = 'https://scholar.google.com/scholar?hl=en&as_sdt=0%2C5&q={}%40{}%09{}&btnG='
        for key in self.keys:
            response = self.get_response(query.format(str(key['email']).split('@')[0], str(key['email']).split('@')[1], str(key['name']).replace(' ', '+')))
            content = response.text
            soup_page= BeautifulSoup(content, 'html.parser')


        pass
    
    def get_detail_data(self):
        pass
    
    def update_csv(self):
        pass

    def read_keys(self, filepath):
        print('Reading keywords for searching...')
        try:
            workbook = load_workbook(filename=filepath)
            worksheet = workbook['Sheet1']
            
            for row in worksheet.iter_rows(values_only=True):
                row_data=dict()
                row_data['email'] = row[0]
                row_data['name'] = row[1]
                self.keys.append(row_data)
            self.keys.pop(0)
        except Exception as e:
            print('There is error: ', e, 'reading keywords.') 

    def read_proxies(self, filepath):
        print('Reading proxies ...')
        try:
            workbook = load_workbook(filepath)
            worksheet = workbook['Sheet1']

            for row in worksheet.iter_rows(values_only=True):
                proxy_url = row[0]
                username = row[1]
                pwd = ""
                proxy_address = username + ':' + pwd + '@' + proxy_url
                self.proxies.append(proxy_address)
            self.proxies.pop(0)
        except Exception as e:
            print('There is error: ', e, 'reading proxies.') 

    def get_response(self, url):
        print('Sending Request with : ', url)
        try:
            i = random.randint(0, len(self.proxies) - 1)

            self.session.proxies = {
                str.format('http://{}', self.proxies[i]),
                str.format('https://{}', self.proxies[i]),
            }

            response = self.session.get(url)
            return response
        except Exception as e:
            print('There is error :', e,'sending request')

def main():
    extracotr = Extractor()
    extracotr.read_proxies(PROXIES)
    extracotr.get_data()
if __name__ == "__main__":
    main()
