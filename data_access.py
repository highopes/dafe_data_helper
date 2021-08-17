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

### Check OpenpyXL Version
### adapt openpyxl function call based on version to avoid deprecation warning message
VERSION_MATCH = re.compile(r'^2\.4')
OPENPYXL_VERSION = pkg_resources.get_distribution("openpyxl").version
if not re.match(VERSION_MATCH, OPENPYXL_VERSION):
    OPENPYXLMODE = "2.5"
else:
    OPENPYXLMODE = "2.4"

# Disable warning message logged by openpyxl
warnings.filterwarnings("ignore")


class Exception_Class(Exception):
    """
    Exception class handling the exception raised by this script
    """

    def fatal(self, msg):
        """
        Prints an error message and aborts program execution
        """
        global EXECUTION_STATUS
        EXECUTION_STATUS = "failed"
        sys.stderr.write(msg + "\n")
        sys.exit(1)

    def warning(self, msg):
        """
        Prints a warning message to stderr
        """
        global EXECUTION_STATUS
        EXECUTION_STATUS = "partial_success"
        sys.stderr.write(msg + "\n")


MyErr = Exception_Class()


class DataAccess(object):
    """
    This class provides methods to read or write data in an Excel file
    """

    def open_xls(self, file_name):
        """ Takes file name as input and return an openpyxl workbook object
        @param file_name : excel source file name
        @return workbook : openpyxl workbook object
        """
        try:
            print("Opening excel file %s\n" % file_name)
            workbook = openpyxl.load_workbook(file_name,
                                              data_only=True)  # data_only = True, means value only without formula
            return workbook
        except IOError:
            MyErr.fatal("IOError Can't Open file %s" % file_name)
        except:
            MyErr.fatal("Undefined error opening excel %s data input" % file_name)

    def save_xls(self, workbook, file_name):
        """ Save changes to Excel file
        """
        try:
            print("Saving excel file %s\n" % file_name)
            workbook.save(file_name)
            print("File %s saved. \n" % file_name)
        except IOError:
            MyErr.fatal("IOError Can't Save file %s" % file_name)
        except:
            MyErr.fatal("Undefined error saving excel %s data" % file_name)

    def read_xls(self, workbook, sheet_name):
        """ Input workbook and a worksheet name, return a generator.
        The generator can save memory.
        """
        print("+---------- Loading Data ----------+\n")
        if OPENPYXLMODE == "2.5":
            active_sheet = workbook[sheet_name]
        else:
            active_sheet = workbook.get_sheet_by_name(sheet_name)

        print("+------------------------------------------+\n")
        return active_sheet.values

    def write_row(self, workbook, sheet_name, row_list):
        """ Input workbook and a worksheet name, return a generator.
        The generator can save memory.
        """
        if OPENPYXLMODE == "2.5":
            active_sheet = workbook[sheet_name]
        else:
            active_sheet = workbook.get_sheet_by_name(sheet_name)
        try:
            active_sheet.append(row_list)
        except:
            MyErr.fatal("Fail to write to the Excel File!")

        return active_sheet.values


def main():
    """
    This main function is just for test.
    """
    da = DataAccess()
    wb = da.open_xls('aci_build_input_data_test.xlsx')

    # ws = wb['tenant']
    # print(ws.cell(1, 1).value)

    # active_gen = da.read_xls(wb, 'tenant')
    # for i in active_gen:
    #     print(i)

    da.write_row(wb, 'tenant', ['hangwe-DAFE002', 'config by DAFE on Aug16,20021'])
    da.save_xls(wb, 'aci_build_input_data_test.xlsx')


if __name__ == '__main__':
    main()
