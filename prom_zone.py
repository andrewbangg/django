import pandas as pd
import lxml.etree as et
import uuid
from lxml.builder import E
import sqlite3 as sq

place_file = r'C:\Users\a.ryzhov\Desktop\Andrew\2022\9_september\prom_zona\registr_prom.xlsx'

conn = sq.connect('prom-baza.db')
cur = conn.cursor()

def read_file(place_file):
    re_file = pd.read_excel(place_file,sheet_name='otvet') #с помощью библиотеки pandas получаем данные с exel
    return re_file

def head_file(read_fil):
    head_files = read_fil.columns.ravel()
    return head_files

def file_dict(file_read):
    fil = read_file(place_file)
    fil_dict = fil.to_dict(orient='index')
    return fil_dict

def structura():
    xsi = 'http://www.w3.org/2001/XMLSchema-instance'
    holding = et.Element('holding',
                      {'{%s}noNamespaceSchemaLocation' % xsi: 'HoldingStructure_2.xsd'},
                      nsmap={'xsi': xsi})
    employees = et.SubElement(holding, 'employees')
    departments = et.SubElement(holding,'departments')
    glob_dep_id = str(uuid.uuid4())
    department = et.SubElement(departments, 'department')
    department.set('id', glob_dep_id)
    department.append(E.title('Производители Урала "Промзона Прорыв"'))
    glob_position = et.SubElement(department, 'positions')
    departments_2 = et.SubElement(department, 'departments')
    for i in file_dict(read_file(place_file)).values():
        employee = et.SubElement(employees, 'employee')
        employee_id = str(uuid.uuid4())
        employee.set('id', employee_id)
        if type(i['ФИО']) != str:
            employee.append(E.fullName(str(i['ФИО'])))
        else:
            employee.append(E.fullName(i['ФИО']))
        employee.append(E.login(str(i['Email'].lower())))
        employee.append(E.email(str(i['Email'].lower())))
        position = et.SubElement(employee, 'position')
        position_id = str(uuid.uuid4())
        position.set('id',position_id)
        employee.append(E.workPhone(str(i['Телефон'])))
        dep = et.SubElement(departments_2,'department')
        department_id = str(uuid.uuid4())
        dep.set('id',department_id)
        dep.append(E.title(i['Название компании']))
        poss = et.SubElement(dep,'positions')
        pos = et.SubElement(poss,'position')
        pos.set('id', position_id)
        pos.append(E.title('Генеральный директор'))
        emp = et.SubElement(pos,'employee')
        emp.set('id',employee_id)

    return holding


hol = '<?xml version="1.0" encoding="UTF-8"?>\n%s' % et \
      .tostring(structura(), pretty_print=True, encoding='unicode') \
      .strip()

with open('get.xml','w',encoding='utf-8') as get:
    get.write(hol)
