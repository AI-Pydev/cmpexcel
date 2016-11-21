import xlrd
import os
import sys


def excel_file(arg):
    if os.path.exists(arg):
        wb = xlrd.open_workbook(arg)
        sheet = wb.sheet_by_index(0)
        rows = sheet.nrows
        cols = sheet.ncols
        excel_content = []
        for row in range(rows):
            for col in range(cols):
                cell = sheet.cell(row, col)
                # print "cell",cell
                excel_content.append(cell.value)
        return excel_content
    else:
        print "File Does Not Exist! Default Folder Path %s" %(arg)


def compare_excel_files(arg1, arg2):
    if len(arg1)==len(arg2):
        print "Hey! Item count are matching in both files"
        #print "arg1",(set(arg1) & set(arg2))

        output = [i for i, j in zip(arg1, arg2) if i != j]
        if not output:
            print "Contents are matching"
        else:
            print "Contents are different"

    else:
        print "Item count are different"

    # print "results are",zip(arg1, arg2)


if __name__=='__main__':
    print 'Hint:- Default Folder Path "%s"' %(os.getcwd())

    if len(sys.argv) == 3:

        if sys.argv[1] and sys.argv[2]:
            # print "Excel file data",excel_file_data(sys.argv[1], sys.argv[2])

            if sys.argv[1].endswith('.xls') or sys.argv[1].endswith('.xlsx'):
                first_excel = sys.argv[1]
                file1_data = excel_file(first_excel)
            else:
                file1_data = None
                print "Error: Invalid File %s only xls and xlsx are allowed" % (sys.argv[1])

            if sys.argv[2].endswith('.xls') or sys.argv[2].endswith('.xlsx'):
                second_excel = sys.argv[2]
                file2_data = excel_file(second_excel)
            else:
                file2_data = None
                print "Error: Invalid File %s only xls and xlsx are allowed" % (sys.argv[1])

            if file1_data and file2_data is not None:
                # print "file 2 data",file2_data
                # print "file 1 data",file1_data
                compare_excel_files(file1_data, file2_data)
    else:
        print "It's Required two excel files Ex: filename.py file1.xlsx file2.xlsx"




