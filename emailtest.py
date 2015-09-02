__author__ = 'bdrummon'

import httplib2
import os
import re

import gspread

def emailtest(gc):
    print "emailtest() called"
    sheet = gc.open("Copy of Q3-15 Capitalized Projects Reporting")
    projects_worksheet = sheet.worksheet("Reporting - Q3-15")

    header_of_projects_table = 10
    projects_records = projects_worksheet.get_all_records(empty2zero=False, head=header_of_projects_table)

    for project in projects_records:
        print "Entered By:", project['Entered By']

    True

if __name__ == '__main__':
    emailtest()

