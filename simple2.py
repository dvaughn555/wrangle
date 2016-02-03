"""
Your task is as follows:
- read the provided Excel file
- find and return the min, max and average values for the COAST region
- find and return the time value for the min and max entries
- the time values should be returned as Python tuples

Please see the test function for the expected return format
"""

import xlrd
from zipfile import ZipFile

datafile = "2013_ERCOT_Hourly_Load_Data.xls"


def open_zip(zipfile):
    with ZipFile('{0}.zip'.format(zipfile), 'r') as myzip:
        myzip.extractall()


def parse_file(thefile):
    workbook = xlrd.open_workbook(thefile)
    sheet = workbook.sheet_by_index(0)

    ### example on how you can get the data
    sheet_data = [[sheet.cell_value(r, col) for col in range(sheet.ncols)] for r in
                  range(sheet.nrows)]

    ### other useful methods:
    print "\nROWS, COLUMNS, and CELLS:"
    print "Number of rows in the sheet:",
    print sheet.nrows
    print "Type of data in cell (row 3, col 2):",
    print sheet.cell_type(3, 2)
    print "Value in cell (row 3, col 2):",
    print sheet.cell_value(3, 2)
    print "Get a slice of values in column 3, from rows 1-3:"
    print sheet.col_values(3, start_rowx=1, end_rowx=4)

    print "\nDATES:"
    print "Type of data in cell (row 1, col 0):",
    print sheet.cell_type(1, 0)
    exceltime = sheet.cell_value(1, 0)
    print "Time in Excel format:",
    print exceltime
    print "Convert time to a Python datetime tuple, from the Excel float:",
    print xlrd.xldate_as_tuple(exceltime, 0)

    maxtime = sheet.cell_value(1, 0)
    mintime = sheet.cell_value(1, 0)
    maxvalue = sheet.cell_value(1, 1)
    minvalue = sheet.cell_value(1, 1)
    totalcoast = sheet.cell_value(1, 1)

    for i in range(2, sheet.nrows):
        row = sheet.row_slice(i)
        totalcoast += row[1].value
        if maxvalue < row[1].value:
            maxtime = row[0].value
            maxvalue = row[1].value
        if minvalue > row[1].value:
            mintime = row[0].value
            minvalue = row[1].value

    data = {
        'maxtime': xlrd.xldate_as_tuple(maxtime, 0),
        'maxvalue': maxvalue,
        'mintime': xlrd.xldate_as_tuple(mintime, 0),
        'minvalue': minvalue,
        'avgcoast': totalcoast / (sheet.nrows - 1)
    }

    print data

    # data = {
    #     'maxtime': (0, 0, 0, 0, 0, 0),
    #     'maxvalue': 0,
    #     'mintime': (0, 0, 0, 0, 0, 0),
    #     'minvalue': 0,
    #     'avgcoast': 0
    # }

    return data


def test():
    open_zip(datafile)
    data = parse_file(datafile)

    assert data['maxtime'] == (2013, 8, 13, 17, 0, 0)
    assert round(data['maxvalue'], 10) == round(18779.02551, 10)
    print "YAY!"


test()
