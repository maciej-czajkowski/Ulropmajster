def naturals(n):  # return naturals series up to n-1
    result = []
    i = 0
    while i < n:
        if i <10:
            temp = "0" + str(i)
            result.append(temp)  # see below
        else:
            result.append(str(i))
        i += 1
    return result

def naturals2(n, m):  # return naturals series from a up to m-1
    result = []
    i = n
    while i < m:
        if i <10:
            temp = "0" + str(i)
            result.append(temp)  # see below
        else:
            result.append(str(i))
        i += 1
    return result

import pandas as pd

def getdrivers():            #creates a list of dictionaries with drivers specs
    file = pd.read_excel('Tabelka_Kierowcy.xlsx', sheet_name='Sheet1')

    drivers = []

    for i in range((len(file.index))):
        drivers.append({'name': '', 'employed': '', 'license': '', 'date_of_birth': ''})
        drivers[i]['name'] = file['Kierowca'][i]
        drivers[i]['employed'] = file['Zatrudniony_od'][i]
        drivers[i]['date_of_birth'] = file['Data_urodzenia'][i]
        drivers[i]['license'] = file['Prawo_Jazdy'][i]
    return drivers

def getemployees():             #creates a list of dictionaries with emplyees specs
    file = pd.read_excel('Tabelka_Kierowcy.xlsx', sheet_name='Sheet2')

    employees = []

    for i in range((len(file.index))):
        employees.append({'name': '', 'position': ''})
        employees[i]['name'] = file['Pracownik'][i]
        employees[i]['position'] = file['Stanowisko'][i]
    return employees
