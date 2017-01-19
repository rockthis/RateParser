import requests
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
from xlsxwriter import Workbook
import xlrd
import os



abba = '54.30K'
bbaab= int(abba.replace('.','').replace('K','0'))
print(bbaab)