#!/usr/bin/env python3 

from bs4 import BeautifulSoup 
import xlsxwriter 
import sys
import datetime
from lxml import etree as et 


DATE = datetime.datetime.today().strftime('%m/%d/%Y')


def child_finder(element, child_name):
    try:
        return element.findChild(child_name).get_text()
    except AttributeError:
        return ''


def get_contributors(con_list):
    auth = con_list[0] 
    authname = auth.findChild('b036').get_text()
    authnamei = auth.findChild('b037').get_text()
    authbio = child_finder(auth, 'b044')
    #authbio = auth.findChild('b044').get_text() 
    add_contribs = []
    for contrib in con_list[1:]:
        addname = contrib.findChild('b036').get_text()
        addnamei = contrib.findChild('b037').get_text()
        addtype = contrib.findChild('b035').get_text()
        add_contribs.append('{} ({}) [{}]'.format(addname, addnamei, addtype))
    format_contribs = ' / '.join(add_contribs)
    listy = [authname, authnamei, authbio, format_contribs]
    return listy


def subject_formater(subject_list, b064):
    if b064 != '':
        return_list = [b064]
    else:
        return_list = []
    keywords = []
    for s in subject_list:
        typ = s.findChild('b067').get_text()
        if typ == '10':
            print(s.findChild('b069'))
            return_list.append(s.findChild('b069').get_text())
        else:
            print(s.findChild('b070'))
            try:
                keywords.append(s.findChild('b070').get_text())
            except AttributeError:
                pass
    try:
        keywords = keywords.pop()
    except IndexError:
        keywords = ''

    return ' '.join(return_list), keywords


def date_formater(pubdate):
    year = pubdate[:4] 
    month = pubdate[4:6]
    day = pubdate[6:]
    return '{}/{}/{}'.format(month, day, year)


def find_page_num(extent):
    for ext in extent:
        if ext.findChild('b218').get_text() == '00':
            return ext.findChild('b219').get_text()


def find_description(othertext):
    for ot in othertext:
        if ot.findChild('d102').get_text() == '03':
            return ot.findChild('d104').get_text()


def product_parser(product, worksheet, row):
    publisher = product.findChild('imprint').findChild('b079').get_text()
    worksheet.write(row, 1, publisher)
    worksheet.write(row, 2, 2)
    isbn = product.findChild('productidentifier').findChild('b244').get_text()
    worksheet.write(row, 3, isbn)
    
    title = product.findChildren('title').pop()
    booktitle = title.findChild('b203').get_text()
    worksheet.write(row, 4, booktitle)
    subtitle = child_finder(title, 'b029')
    worksheet.write(row, 5, subtitle) 

    series = product.findChild('series')
    seriestitle = child_finder(series, 'b203')
    
    worksheet.write(row, 6, seriestitle)

    authnames = product.findChildren('contributor')
    contribs = get_contributors(authnames)
    i = 0 
    for col in range(8,12):
        worksheet.write(row, col, contribs[i])
        i += 1 
    b064 = child_finder(product, 'b064')
    subjects = product.findChildren('subject')
    subj_insert, keywords = subject_formater(subjects, b064)
    worksheet.write(row, 12, subj_insert)

    keywords = keywords.replace(';', ' / ')
    worksheet.write(row, 13, keywords)

    language = product.findChild('language').findChild('b252').get_text()
    if language == 'eng':
        language = 'English'
    worksheet.write(row, 14, language)

    pubdate = product.findChild('b003').get_text() 
    inputdate = date_formater(pubdate) 
    worksheet.write(row, 15, inputdate)
    worksheet.write(row, 16, inputdate)

    worksheet.write(row, 17, DATE)

    extent = product.findChildren('extent')
    pagenum = find_page_num(extent)
    worksheet.write(row, 18, pagenum)

    othertext = product.findChildren('othertext')
    description = find_description(othertext)
    worksheet.write(row, 19, description)

    pricing = product.findChild('price').findChild('j151').get_text()
    price = 'WW ' + pricing 
    worksheet.write(row, 28, price)
    worksheet.write(row, 29, price)

    
def main():
    xmlfile = sys.argv[1]
    workbook = xlsxwriter.Workbook('INScript.xlsx') 
    worksheet = workbook.add_worksheet() 
    with open(xmlfile, encoding='utf8') as fp:
        xml = BeautifulSoup(fp, features="lxml")
    products = xml.find_all('product')
    row = 1
    for product in products:
        product_parser(product, worksheet, row)
        row += 1
    workbook.close()


if __name__ == '__main__':
    main()
