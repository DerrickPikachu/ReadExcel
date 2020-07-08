import xlrd


def main():
    excel = xlrd.open_workbook('test.xlsx')

    # sheets = excel.get_sheets()
    # print(sheets)
    sheet = excel.sheet_by_index(0)

    rowNum = sheet.nrows
    colNum = sheet.ncols
    print("We get {} rows and {} cols".format(rowNum, colNum))

    row3 = sheet.row_values(2)
    row5 = sheet.row_values(4)
    print("row 3: {}".format(row3))
    print("row 5: {}".format(row5))

    col1 = sheet.col_values(0)
    print("row 3 and col 1 data is: {}".format(col1[2]))


if __name__ == "__main__":
    main()