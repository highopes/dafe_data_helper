#!/usr/bin/env python
###################################################################################
#                           Written by Wei, Hang                                  #
#                          weihang_hank@gmail.com                                 #
###################################################################################
"""
This module is used to read or write the specified contents of an Excel file.
A small amount of code references the DAFE project.
"""

import argparse
import os
import warnings
import sys
import time
import re
from xml.dom import minidom
import openpyxl
import yaml
import pkg_resources
import data_access

EXCEL_FILE = "aci_build_input_data_test.xlsx"


def Substract_Lists(lst1, lst2):
    """
    Remove elements in lst2 from lst1
    """
    return list(filter(lambda x: x not in lst2, lst1))


def Tenant_Data():
    """
    This function prepare data for configuring Tenant
    """
    # # Please modify following as needed
    # data_list is the row to be append to the Excel file
    # for Tenant the list is: name, description, name_alias, security_domain, status
    # set the initial value as needed
    data_list = ["", "Config by DAFE data helper", "", "", ""]
    # Tenant Name Format is: Tenant{INDEX}{KEYWORD1}{KEYWORD2}
    INDX_INCLUDE = [*range(1, 23 + 1)]  # elements included, *range(1, 23+1) means 1-23
    INDX_EXCLUDE = [2, *range(4, 6 + 1), 9]  # elements excluded
    INDEX = Substract_Lists(INDX_INCLUDE, INDX_EXCLUDE)
    KEYWORD1 = "_hangwe"
    KEYWORD2 = "_DAFE"
    # Following is the EXCEL sheet name
    SHEETNAME = "tenant"

    # Prepare data file to be filled in
    dataaccess = data_access.DataAccess()
    workbook = dataaccess.open_xls(EXCEL_FILE)

    # Precess the dynamic elements in data_list and write to the Excel file
    for indx in INDEX:
        name = "Tenant{}{}{}".format("%03d" % indx, KEYWORD1, KEYWORD2)
        data_list[0] = name

        dataaccess.write_row(workbook, SHEETNAME, data_list)

    # Save to Excel File
    dataaccess.save_xls(workbook, EXCEL_FILE)


def main():
    """
    This main function is just for test.
    """
    Tenant_Data()


if __name__ == '__main__':
    main()
