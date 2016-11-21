import xlrd
import os
import sys


def excel_file_data(arg1, arg2):
    if os.path.exists(arg1) and os.path.exists(arg2):
        wb1 = xlrd.open_workbook(arg1)
        wb2 = xlrd.open_workbook(arg2)
        sheet1 = wb1.sheet_by_index(0)
        sheet2 = wb2.sheet_by_index(0)
        if sheet1.nrows == sheet2.nrows and sheet1.ncols == sheet2.ncols:
            print "Items count are matching"
            rows = sheet1.nrows
            cols = sheet1.ncols
            unmatched = []
            for row in range(rows):
                for col in range(cols):
                    # print sheet1.cell(row, col).value,'==',sheet2.cell(row, col).value
                    if sheet1.cell(row, col).value != sheet2.cell(row, col).value:
                        print "Excel file row %s and column %s data are not matching" % (row, col)
                        unmatched.append(sheet1.cell(row, col))
            if not unmatched:
                print "Contents are matching"
            else:
                print "Contents are different"
        else:
            print "Item count are different"
    else:
        print "Invalid Path, Default Folder Path %s" %(os.getcwd())


if __name__ == '__main__':
    print 'Hint:- Default Folder Path "%s"' %(os.getcwd())

    if len(sys.argv) == 3:

        if sys.argv[1] and sys.argv[2]:
            # print "Excel file data",excel_file_data(sys.argv[1], sys.argv[2])

            if sys.argv[1].endswith('.xls') or sys.argv[1].endswith('.xlsx'):
                file1_data = sys.argv[1]
            else:
                file1_data = None
                print "Error: Invalid File %s only xls and xlsx are allowed" % (sys.argv[1])

            if sys.argv[2].endswith('.xls') or sys.argv[2].endswith('.xlsx'):
                file2_data = sys.argv[2]
            else:
                file2_data = None
                print "Error: Invalid File %s only xls and xlsx are allowed" % (sys.argv[1])

            if file1_data and file2_data is not None:
                # print "file 2 data",file2_data
                # print "file 1 data",file1_data
                excel_file_data(sys.argv[1], sys.argv[2])


    else:
        print "It's Required two excel files Ex: filename.py file1.xlsx file2.xlsx"




