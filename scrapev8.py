import pandas as pd # library for data analysis
import requests # library to handle requests
from bs4 import BeautifulSoup # library to parse HTML documents
import xlrd
from xlutils.copy import copy

global sing, plu, prat, perf, ich, du, er, kom, sup

def searchName(name):
    global sing, plu
    # get the response in the form of html
    name = name[0].upper() + name[1:] 
    wikiurl="https://de.wiktionary.org/wiki/" + name
    response=requests.get(wikiurl)
    # parse data from the html into a beautifulsoup object
    soup = BeautifulSoup(response.text, 'html.parser')
    indiatable=soup.find('table',{'class':"wikitable"})
    if indiatable != None:
        df=pd.read_html(str(indiatable), encoding='utf-8')
        # convert list to dataframe
        df=pd.DataFrame(df[0])
        if (df.columns.values[1] == 1) and (df.columns.values[2] == 2):
            print("test1")
            sing = df[1][2]
            plu = df[2][2]
        elif (df.columns.values[1] == 'Singular') and (df.columns.values[2] == 'Plural'):
            print("test2")
            sing = df['Singular'][0]
            plu = df['Plural'][0]
        elif (df.columns.values[1] == 'Singular 1') and (df.columns.values[2] == 'Singular 2') and (df.columns.values[3] == 'Plural'):
            print("test3")
            sing = df['Singular 1'][0]
            plu = df['Plural'][0]
        elif (df.columns.values[1] == 'Singular') and (df.columns.values[2] == 'Plural 1') and (df.columns.values[3] == 'Plural 2'):
            print("test4")
            sing = df['Singular'][0]
            plu = df['Plural 1'][0]
        else:
            print("error1")
            sing = 'error1'
            plu = 'error1'
    else:
        print("error2")
        sing = 'error2'
        plu = 'error2'

def searchVerb(verb):
    global prat, perf, ich, du, er
    # get the response in the form of html
    verb = verb[0].lower() + verb[1:] 
    wikiurl="https://de.wiktionary.org/wiki/" + verb
    response=requests.get(wikiurl)
    # parse data from the html into a beautifulsoup object
    soup = BeautifulSoup(response.text, 'html.parser')
    indiatable=soup.find('table',{'class':"wikitable"})
    #print(indiatable)
    if indiatable != None:
        df=pd.read_html(str(indiatable), encoding='utf-8')
        #print(df)
        # convert list to dataframe
        df=pd.DataFrame(df[0])
        if (df.columns.values[1] == 'Person') and (df.columns.values[2] == 'Wortform'):
            print("test2")
            if df['Wortform.2'][8] == "haben" :
                prat = df['Wortform'][3]
                perf = ("hat" + " " + df['Wortform'][8])
                ich = (df['Person'][0] + " " + df['Wortform'][0])
                du = (df['Person'][1] + " " + df['Wortform'][1])
                er = ("er" + " " + df['Wortform'][2])
            elif df['Wortform.2'][8] == "sein" :
                prat = df['Wortform'][3]
                perf = ("ist" + " " + df['Wortform'][8])
                ich = (df['Person'][0] + " " + df['Wortform'][0])
                du = (df['Person'][1] + " " + df['Wortform'][1])
                er = ("er" + " " + df['Wortform'][2])
            elif df['Wortform.2'][8] == "sein, haben" :
                prat = df['Wortform'][3]
                perf = ("ist, hat" + " " + df['Wortform'][8])
                ich = (df['Person'][0] + " " + df['Wortform'][0])
                du = (df['Person'][1] + " " + df['Wortform'][1])
                er = ("er" + " " + df['Wortform'][2])
            elif df['Wortform.2'][8] == "haben, sein" :
                prat = df['Wortform'][3]
                perf = ("hat, ist" + " " + df['Wortform'][8])
                ich = (df['Person'][0] + " " + df['Wortform'][0])
                du = (df['Person'][1] + " " + df['Wortform'][1])
                er = ("er" + " " + df['Wortform'][2])
        else:
            print("error1")
            prat = 'error1'
            perf = 'error1'
            ich = 'error1'
            du = 'error1'
            er = 'error1'
    else:
        print("error2")
        prat = 'error2'
        perf = 'error2'
        ich = 'error2'
        du = 'error2'
        er = 'error2'

def searchOther(other):
    global kom, sup
    # get the response in the form of html
    other = other[0].lower() + other[1:] 
    wikiurl="https://de.wiktionary.org/wiki/" + other
    response=requests.get(wikiurl)
    # parse data from the html into a beautifulsoup object
    soup = BeautifulSoup(response.text, 'html.parser')
    indiatable=soup.find('table',{'class':"wikitable"})
    #print(indiatable)
    if indiatable != None:
        df=pd.read_html(str(indiatable), encoding='utf-8')
        #print(df)
        
        # convert list to dataframe
        df=pd.DataFrame(df[0])
        #print(df.columns.values)
        if (df.columns.values[0] == 'Deklination der Kardinalzahlen 2–12'):
            print("test1")
            kom = ""
            sup = ""
        elif (df.columns.values[0][0] == 'attributiv (vor Substantiv)'):
            print("test2")
            kom = ""
            sup = ""
        elif (df.columns.values[1] == 'Komparativ') and (df.columns.values[2] == 'Superlativ'):
            print("test3")
            if df['Komparativ'][0] == "—" and df['Superlativ'][0]== "—" :
                kom = ""
                sup = ""
            else:
                kom = df['Komparativ'][0]
                sup = df['Superlativ'][0]
        elif (df.columns.values[1] == 'Singular') and (df.columns.values[2] == 'Plural'):
            print("test4")
            kom = ""
            sup = ""
        elif (df.columns.values[1][0] == 'Singular') and (df.columns.values[4][0] == 'Plural'):
            print("test5")
            kom = ""
            sup = ""
        elif (df.columns.values[0] == 'Singular') and (df.columns.values[1] == 'Plural'):
            print("test6")
            kom = ""
            sup = ""
        else:
            print("error1")
            kom = 'error1'
            sup = 'error1'
    else:
        print("noTable")
        kom = ''
        sup = ''

def translateTr(other):
    # get the response in the form of html
    wikiurl="https://tr.pons.com/çeviri/almanca-türkçe/" + other
    response=requests.get(wikiurl)
    # parse data from the html into a beautifulsoup object
    soup = BeautifulSoup(response.text, 'html.parser')
    entry=soup.find('div',{'class':'entry first', 'rel':other})
    #print(entry)
    if entry != None:
        translations=entry.find('div',{'class':'translations first'})
        #print(translations)
        source = translations.find('div',{'class':'target'})
        #print(source)
        tr = source.find_all('a')
        ret = ""
        for i in tr:
            ret = ret + i.text + " "
        ret = ret.rstrip()
        return ret
    else:
        return "errorTr"

def translateEn(other):
    # get the response in the form of html
    wikiurl="https://tr.pons.com/çeviri/almanca-ingilizce/" + other
    response=requests.get(wikiurl)
    # parse data from the html into a beautifulsoup object
    soup = BeautifulSoup(response.text, 'html.parser')
    entry=soup.find('div',{'class':'entry first', 'rel':other})
    #print(entry)
    if entry != None:
        translations=entry.find('div',{'class':'translations first'})
        #print(translations)
        source = translations.find('div',{'class':'target'})
        #print(source)
        en = source.find_all('a')
        ret = ""
        for i in en:
            ret = ret + i.text + " "
        ret = ret.rstrip()
        return ret
    else:
        return "errorEn"
    

#print(translateTr("Handy"))
#searchOther("mein")
#print(kom, sup)

# Give the location of the file
path = ("kelimeler.xls")

rb = xlrd.open_workbook(path)
rs = rb.sheet_by_index(0)
wb = copy(rb)
ws = wb.get_sheet(0)

_list = rs.col_values(0)

for i in range(rs.nrows):
    name = _list[i]
    name = name[0].upper() + name[1:] 
    searchName(name)
    turkce = translateTr(name)
    english = translateEn(name)
    print(i, name, sing, plu, turkce, english)
    ws.write(i, 1, sing)
    ws.write(i, 2, plu)
    ws.write(i, 3, turkce)
    ws.write(i, 4, english)

wb.save(path)
rb = xlrd.open_workbook(path)
rs = rb.sheet_by_index(1)
wb = copy(rb)
ws = wb.get_sheet(1)

_list = rs.col_values(0)

for i in range(rs.nrows):
    verb = _list[i]
    verb = verb[0].lower() + verb[1:] 
    searchVerb(verb)
    turkce = translateTr(verb)
    english = translateEn(verb)
    print(i, verb, prat, perf, ich, du, er, turkce, english)
    ws.write(i, 1, prat)
    ws.write(i, 2, perf)
    ws.write(i, 3, ich)
    ws.write(i, 4, du)
    ws.write(i, 5, er)
    ws.write(i, 6, turkce)
    ws.write(i, 7, english)

wb.save(path)

rb = xlrd.open_workbook(path)
rs = rb.sheet_by_index(2)
wb = copy(rb)
ws = wb.get_sheet(2)

_list = rs.col_values(0)

for i in range(rs.nrows):
    other = _list[i]
    other = other[0].lower() + other[1:] 
    searchOther(other)
    turkce = translateTr(other)
    english = translateEn(other)
    print(i, other, kom, sup, turkce, english)
    ws.write(i, 1, kom)
    ws.write(i, 2, sup)
    ws.write(i, 3, turkce)
    ws.write(i, 4, english)

wb.save(path)
#'''