import requests
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
from xlsxwriter import Workbook
import xlrd
import os

workbook = Workbook('forex_200emp.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0,0,'SimilarWeb')
worksheet.write(0,1,'CompanyName')
worksheet.write(0,2,'LinkedinLink')
worksheet.write(0,3, 'CompanySite')

pathmac = os.path.expanduser('~/Desktop/Payzoff/linkedin/new_files/Forex/forex_finance_more200epm_170/')
path = 'C:\\Users\\Admin\\Desktop\\Linkedin_files\\Forex\\forex_finance_200empl_324q\\'

filename = []
finalpath = []

for i in range (len(os.listdir(path))):                 #Выбираем файлы в директории windows
    filename.append(os.listdir(path)[i])
    finalpath.append(path + str(filename[i]))

# for i in range (len(os.listdir(pathmac))):                 #Выбираем файлы в директории mac
#     filename.append(os.listdir(pathmac)[i])
#     finalpath.append(pathmac + str(filename[i]))

def read_file(path):                                     #Функция чтения файла
    file = open(path, encoding="utf8")
    data = file.read()
    file.close()
    return data

linkedin_link = []
company_names = []

for i in range (len(finalpath)):
                                                        #Парсим все файлы которые есть в директории,
    pathtofile = str(finalpath[i])                      # и наполняем значениями inkedin_link, company_names
    soup = BeautifulSoup(read_file(pathtofile), 'lxml')

    for a in soup.find_all('a', {'class': 'title main-headline'}):
        linkedin_link.append(a['href'].rsplit('?', 1)[0])

    for a in soup.find_all('a', {'class': 'title main-headline'}):
        company_names.append(a.getText())


duplicates = []                                             #проверка на уникальность списка
for value in company_names:
    if company_names.count(value) > 1:
        if value not in duplicates:
            duplicates.append(value)


for el in duplicates:
    indices = [i for i, x in enumerate(company_names) if x == el]
    del indices[0]
    for j in indices[::-1]:
        del company_names[j]
        del linkedin_link[j]


def unique_list(l):                                     #проверка на уникальность списка
    ulist = []
    [ulist.append(x) for x in l if x not in ulist]
    return ulist

linkedin_link = unique_list(linkedin_link)
company_names = unique_list(company_names)




for i in range(len(linkedin_link)):                     #Пишем в файл результаты
    worksheet.write(1 + i, 1, company_names[i])
    worksheet.write(1 + i, 2, linkedin_link[i])


workbook.close()
    # outputFile = "parsedTable.html"
    # f = open(outputFile, 'w')
    # f.write(soup.prettify())
    # {'valign': re.compile('top')}

    # workbook = xlrd.open_workbook('forex_10emp_900.xlsx')
    # worksheet = workbook.sheet_by_index(0)
    # rows = worksheet.nrows
    # print(rows)
    # workbook.close()
