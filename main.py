from openpyxl import load_workbook
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import lxml
from configparser import ConfigParser

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import pandas as pd
import openpyxl
import os, sys


class CourseName:
    def __init__(self, html_table):
        self.html_table = html_table

    def pull_course_name(self):
        # for i in range(len(self.html_table)):
        # print('pull course table', self.html_table)
        # print(h2_source[table_count].text.strip())
        # print('table count', table_count)
        course_name = h2_source[table_count].text.strip()
        course_name = course_name.split()
        course_name = course_name[0] + ' ' + course_name[1]
        # print('course', course_name)
        department = course_name.split()
        department = department[0]
        department = department.split()
        department = department[0]
        # print('dept', department)
        return course_name, department


class SessionName:
    row_count = 0

    def __init__(self, html_table):
        self.html_table = html_table
        # print('html table', html_table)

    def pull_session(self):
        rows = self.html_table.find_all('tr')
        # print('tr rows', rows)

        for row in rows:
            # print('row', row)
            td_row = row.find_all('td')
            # print('td_row', td_row)
            cols = [x.text.strip() for x in td_row]
            # print('cols', cols)
        if len(cols) == 2:
            for item in cols:
                if 'Session' in item:
                    session = item
                    # print('session', session)
                    return session


class TableWork:
    length = 0

    # def __init__(self, html_table, course_name, session):
    #     self.html_table = html_table
    #     self.course_name = course_name
    #     self.session = session
    def __init__(self, html_table, course_name, department):
        self.html_table = html_table
        self.course_name = course_name
        self.deparment = department

    def extract_row(self):
        cols = []
        session_row = self.html_table.find_all('td', {'colspan': '14'})
        session_row2 = self.html_table.find_all('td', {'class': 'sess2head'})
        rows = self.html_table.find_all('tr')
        for row in rows:
            # print('row', row)
            # print(len(row))
            if len(row) == 3:
                cols1 = row.find_all('td')
                session = [x.text.strip() for x in cols1]
                # print('cols2', session)
                continue
            elif len(row) == 33:
                cols = row.find_all('td')
                cols = [x.text.strip() for x in cols]
                # print('cols', cols)

            else:
                continue
            # cols.insert(0, self.course_name)
            # print(session)
            cols.insert(0, self.deparment)
            cols.insert(1, self.course_name)
            cols.insert(2, session[1])
            # print(cols)
            enrollment_df.loc[TableWork.length] = cols
            TableWork.length += 1
            # print(enrollment_df)
            # for item in row:
            #     print('item', item)
            #     print(len(item))
        # courserows = self.html_table.find_all('tr', {'bgcolor':'lightgrey'})
        # print ('session1', session_row)
        # print('session2', session_row2)
        # print('rows', rows)
        # print(len(item))
        # print(len(rows))
        # session = ""
        # for row in rows:
        #     # print(row)
        #     for item in row:
        #         # print(item)
        #         if 'tr bgcolor' in item:
        #             td = row.find_all('td')
        #             cols = [x.text.strip() for x in td]
        #             if len(cols) == 17:
        #                 cols.insert(0, self.course_name)
        #                 enrollment_df.loc[TableWork.length] = cols
        #
        #     else:
        #         cols = row.find_all('td')
        #         cols = [x.text.strip() for x in cols]
        #         if len(cols) == 2:
        #             for item in cols:
        #                 if 'Session' in item:
        #                     session = cols[1]
        #                     break
        # for row in rows:
        #     # print('row', row)
        #     cols = row.find_all('td')
        #     cols = [x.text.strip() for x in cols]
        #     if len(cols) == 17:
        #         cols.insert(0, self.course_name)
        #         cols.insert(1, session)
        #         enrollment_df.loc[TableWork.length] = cols
        #         TableWork.length += 1
        # print('cols', cols)
        # print(enrollment_df)

        # if len(cols) == 2:
        #     print('cols', cols[1])
        #     # enrollment_df.loc[SessionName.row_count, 'Session'] = cols[1]
        #     print('row count', SessionName.row_count)
        #     # enrollment_df.loc[table_count, 'Session'] = cols[1]
        #     SessionName.row_count += 1


class GroupDepartments:
    headers = ['Dept', 'Course', 'Session', 'Class', 'Start', 'End', 'Days', 'Room', 'Size', 'Max', 'Wait', 'Cap',
               'Seats', 'WaitAv', 'Status', 'Instructor', 'Type', 'Hours', 'Books', 'Modality']
    df = pd.DataFrame()
    pd.set_option('display.max_columns', None)

    def __init__(self, department, df):
        self.department = department
        self.df = df

    def group_by_departments(self):
        grp = self.df.get_group(department)
        grp.to_excel(department + '/' + department + '.xlsx')
        grp_courses = grp.groupby(['Course'])['Size', 'Max'].agg(['count', 'sum'])
        mod_grp_courses = grp.groupby(['Modality'])['Size', 'Max'].agg(['count', 'sum', 'mean'])
        grp_courses.to_excel(department + '/' + department + 'ttl.xlsx')
        mod_grp_courses.to_excel(department + '/' + department + '_modalities.xlsx')

s = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=s)
# driver = webdriver.Chrome("C:/Users/family/PycharmProjects/chromedriver.exe")
driver.get('https://secure.cerritos.edu/schedule/')
# check_fall_or_spring = driver.find_element(By.XPATH, '/html/body/form/p/b/b/label[1]/input').click()
# check_fall_or_spring = driver.find_element(By.XPATH, '/html/body/b/b/form/p[1]/label/input').click()

# check_fall_or_spring = driver.find_element(By.XPATH, '/html/body/form/p/b/b/label[2]/input').click()
check_all = driver.find_element(By.XPATH, '/html/body/form/table[1]/tbody/tr[1]/td[1]/label/input').click()
# /html/body/form/table[1]/tbody/tr[1]/td[1]/label/input
check_LA = driver.find_element(By.XPATH, '/html/body/form/table[6]/tbody/tr[5]/td[2]/label/input').click()
# /html/body/form/table[6]/tbody/tr[5]/td[2]/label/input


click_View = driver.find_element(By.XPATH, '/html/body/form/p[4]/input').click()
# /html/body/form/p[4]/input
page_loading = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ASL110descs')))
headers = ['Dept', 'Course', 'Session', 'Class', 'Start', 'End', 'Days', 'Room', 'Size', 'Max', 'Wait', 'Cap', 'Seats',
           'WaitAv', 'Status', 'Instructor', 'Type', 'Hours', 'Books', 'Modality']
pd.set_option('display.max_columns', None)
enrollment_df = pd.DataFrame(columns=headers)
# enrollment_df = enrollment_df[['Size', 'Max']].apply(pd.to_numeric)
page_source = driver.page_source
soup = BeautifulSoup(page_source, 'lxml')
# div_table = soup.find_all('div', {'name': 'desc'}).decompose()
# print(div_table)
table_count = 0
h2_source = soup.find_all('h2')
session_source = soup.find_all('tr', {'class': 'sess1head', 'colspan': '14'})
tables = soup.find_all(['table', {'cellspacing': '0', 'class': 'class'}])

for table in tables:
    # print('table', table.prettify())
    c = CourseName(html_table=h2_source)
    course_name, department = c.pull_course_name()
    t = TableWork(html_table=table, course_name=course_name, department=department)
    t.extract_row()
    table_count += 1

enrollment_df['Size'] = pd.to_numeric(enrollment_df['Size'], errors='coerce').fillna(0).astype('int')
enrollment_df['Max'] = pd.to_numeric(enrollment_df['Max'], errors='coerce').fillna(0).astype('int')
enrollment_df['Hours'] = pd.to_numeric(enrollment_df['Hours'], errors='coerce').fillna(0).astype('int')

lecture_df = enrollment_df['Type'] == 'Lecture'
lecture_enrollment_df = enrollment_df[lecture_df]

lecture_enrollment_df.loc[lecture_enrollment_df['Room'] == 'REMOTE', 'Modality'] = 'Remote Course'
lecture_enrollment_df.loc[lecture_enrollment_df['Room'] == 'ONLINE', 'Modality'] = 'Online Course'
rooms = ['LA106', 'LA109', 'LA201', 'SS207', 'LA213', 'LC218', 'LA110', 'SS211', 'SS225', 'LA103', 'SS224', 'LC217',
         'LA211', 'LA202']
for room in rooms:
    lecture_enrollment_df.loc[lecture_enrollment_df['Room'] == room, 'Modality'] = 'Hybrid Course'
lecture_enrollment_df.loc[lecture_enrollment_df['Max'] == 15, 'Modality'] = 'In Person'
df_groupby_dept = lecture_enrollment_df.groupby(['Dept'])


lecture_enrollment_df['FTES'] = lecture_enrollment_df['Size'] * (((lecture_enrollment_df['Hours']/18)*17.5)/525)
lecture_enrollment_df.to_excel('Division_Enrollment.xlsx')

departments = ['AFRS', 'ASL', 'CHIN', 'COMM', 'ENGL', 'ESL', 'FREN', 'GERM', 'JAPN', 'READ', 'SPAN']
for department in departments:
    g = GroupDepartments(department=department, df=df_groupby_dept)
    g.group_by_departments()

