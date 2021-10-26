import pandas as pd
import xlrd
import glob
import re

#test

emails = []
emails2 = []

archivos = glob.glob('archivos/*.xls*')
#print (archivos)

for idz,ittemm in enumerate(archivos):

    xls = xlrd.open_workbook(archivos[idz], on_demand=True)

    hojas = xls.sheet_names()
    for idy,ittem in enumerate(hojas):
        hoja = pd.read_excel(archivos[idz],sheet_name=hojas[idy],encoding="utf-8")

        df = hoja.dropna(axis='columns', how='all')


        for idx, item in enumerate(df.columns):
            for a in df.iloc[:,idx ]:
                if type(a)==str:
                    if a.find("@")>0:
                        if a.find(";")>0:
                            e = a.split(";")
                            for i in e:
                                emails.append(i.lower())
                        else:
                            emails.append(a.replace(u'\xa0', u' ').lower())




for e in emails:
    match = re.findall(r'[\w\.-]+@[\w\.-]+',e)
    for a in match:
        emails2.append(a)
     
emails2 = list(set(emails2))
dff = pd.DataFrame(emails2, columns=['emails'])

dff.to_csv('emails.csv', index=False)