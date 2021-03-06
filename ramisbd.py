# -*- coding: utf-8 -*-
import pandas as pd
import os
from tkinter import filedialog
#import dill
import numpy as np

def p(aaarg):
    print(aaarg)
def pdopenfail(path=""):
    # df = pd.read_csv('apple.csv', index_col='Date', parse_dates=True)
    if path == "":
        path = filedialog.askopenfilename() #  open in explorer
    if path.find("//")>0: #  if it full path
        print("found full path")
        print(path)
    else:
        print("guess path") #  if it not full path maybe it here or desktop
        print(os.getcwd()) #  https://ru.stackoverflow.com/questions/535318/%D0%A2%D0%B5%D0%BA%D1%83%D1%89%D0%B0%D1%8F-%D0%B4%D0%B8%D1%80%D0%B5%D0%BA%D1%82%D0%BE%D1%80%D0%B8%D1%8F-%D0%B2-python
        if os.path.exists(os.getcwd() + "/" + path):
            print("found here")
            path=os.getcwd() + "/" + path
        if os.path.exists(os.environ['USERPROFILE'] + '\Desktop/' + path):
            print(os.environ['USERPROFILE'] + '\Desktop/' + path)
            path = os.environ['USERPROFILE'] + '\Desktop/' + path
    if path[len(path)-3:] == "csv":
        return pd.read_csv(path)
    if path[len(path) - 4:] == "xlsx" or path[len(path) - 3:] == "xls" or path[len(path) - 4:] == "xlsm":
        return pd.read_excel(path)
# def snl_session(fn="savedsession.pkl"):
#     if os.path.exists(fn):  # os.path.isfile() if needed to check this folder or file
#         dill.load_session(fn)
#         print('session loaded')
#     else:
#         dill.dump_session(fn)
#         print('session saved')

vf = pdopenfail()

#snl_session()
res = np.argwhere(vf.values == 'Регион')
lc = 0
rc = 1
if res[0][1] > res[1][1]:
    lc = 1  # на случай смещения второй таблицы
    rc = 0
ltable = vf.iloc[res[lc][0]:, res[lc][1]:res[rc][1]].dropna(axis=0, how='all') # убираем строки до региона, убираем вторую таблицу, удаляем пустые строки
ltable = ltable.dropna(axis=1, how='all')  # убираем пустые столбцы
ltable.columns = ltable.iloc[0]  # в заголовок первую теперь уже строку
ltable=ltable.iloc[1:]  # удалить строку переехавшую в заголовок
rtable = vf.iloc[res[rc][0]:, res[rc][1]:].dropna(axis=0, how='all')
rtable = rtable.dropna(axis=1, how='all')
rtable.columns = rtable.iloc[0]
rtable = rtable.iloc[1:]
rtable = rtable.drop(['Регион', 'Код МО', 'Наименование структурного подразделения медицинской организации', 'Название проекта по улучшению'], axis=1)  # столбцы поп которым идет задвоение
ltable.reset_index(drop=True, inplace=True)
rtable.reset_index(drop=True, inplace=True)
ltable = ltable.join(rtable)
ltable['Регион'] = ltable['Регион'].fillna(1)  # хз как зацепиться за НаН, пришлось менять на 1, чтобы продублировать название региона
for i in range(len(ltable)):
    if ltable.loc[i]['Регион'] == 1:
        ltable.loc[i]['Регион'] = ltable.loc[i-1]['Регион']
ltable = ltable.loc[ltable['Регион'] != "ФМБА России"]
tek1s = ['Регион', 'Код МО', "Наименование структурного подразделения медицинской организации", "Обслуживаемое население", "Количество работников, устроенных по основному месту работы в структурном подразделении медицинской организации, на начало календарного года", "Количество работников, из числа устроенных по основному месту работы в структурном подразделении медицинской организации, обученных методам и инструментам бережливого производства (наличие документа о повышении квалификации), на момент подачи отчета"]
Tek1 = ltable[tek1s] #, "Код МО", "Наименование структурного подразделения медицинской организации", "Обслуживаемое население", "Количество работников, устроенных по основному месту работы в структурном подразделении медицинской организации, на начало календарного года", "Количество работников, из числа устроенных по основному месту работы в структурном подразделении медицинской организации, обученных методам и инструментам бережливого производства (наличие документа о повышении квалификации), на момент подачи отчета"]
Tek2 = Tek1.fillna("нет данных1")
Tek1 = Tek1.groupby(tek1s).size().reset_index(name=('дублей'))
Tek1 = Tek1.replace({"нет данных1": ""})

#snl_session("tek1.pkl")
tek2s=["Код МО", "Наименование структурного подразделения медицинской организации", "Название проекта по улучшению", "Улучшаемый процесс", "Улучшаемый процесс (другое)", "Улучшаемый подпроцесс", "Улучшаемый подпроцесс (другое)", "Дата открытия проекта", "Плановая дата закрытия проекта (в соответствии с паспортом проекта по улучшению)", "Фактическая дата закрытия проекта", " Этап  реализации проекта"]
Tek2 = ltable[tek2s]
Tek2.loc[:, ["Срок реализации проекта", "Количество разработанных стандартов", "Соответствие даты закрытия плановой дате", "Дата, предшествующая отчетному периоду"]] = ""
Tek2 = Tek2.join(ltable[["Полученный экономический эффект", "Связь проекта с организацией медицинской помощи в условиях COVID 19", "Руководителя заместителя руководителя медицинской организации", "Работника РЦ ПМСП"]])
Tek2["Ищем прочих в базе"] = ""
Tek2 = Tek2.join(ltable["Обслуживаемое население"])
tek2s.extend(("Полученный экономический эффект", "Связь проекта с организацией медицинской помощи в условиях COVID 19", "Руководителя заместителя руководителя медицинской организации", "Работника РЦ ПМСП", "Обслуживаемое население"))
Tek2 = Tek2.fillna("нет данных1")
Tek2 = Tek2.groupby(list(Tek2.columns)).size().reset_index(name=('дублей'))
Tek2 = Tek2.replace({"нет данных1": ""})

#snl_session("tek2.pkl")
tek3s=["Код МО", "Название проекта по улучшению", "Улучшаемый процесс", "Улучшаемый процесс (другое)", "Улучшаемый подпроцесс", "Улучшаемый подпроцесс (другое)", "Дата открытия проекта", "Плановая дата закрытия проекта (в соответствии с паспортом проекта по улучшению)", "Фактическая дата закрытия проекта", "Форма стандарта", "Форма стандарта (другое)", "Наименование стандарта"]
Tek3 = ltable[tek3s]
Tek3 = Tek3.fillna("нет данных1")
Tek3 = Tek3.groupby(list(Tek3.columns)).size().reset_index(name=('дублей'))
Tek3 = Tek3.replace({"нет данных1": ""})

with pd.ExcelWriter('Итог.xlsx') as writer:
    Tek1.to_excel(writer, sheet_name='Текущ 1', index=False)
    Tek2.to_excel(writer, sheet_name='Текущ 2', index=False)
    Tek3.to_excel(writer, sheet_name='Текущ 3', index=False)
ltable.to_excel('Исх без фмба.xlsx', sheet_name='прикольно получилось', index=False)
p('типа всё')
sec = input()
