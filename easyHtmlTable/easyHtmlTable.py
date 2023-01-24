"""
    Name: easyHtmlTable
    Author: Ming Yin Kenneth Chau
    Description: Create easy html table from excel file
"""

import openpyxl

# Define the easyHtmlTable function
def easyHtmlTable(filename: str, sheetname: str, containHyper:int = None, firstColumn = 1, firstRow = 1, headerRow = True):
    workBook = openpyxl.load_workbook(filename)
    workSheet = workBook[sheetname]

    # Now we create a list for headers and a list of list for the body
    header = []
    body = []

    # Get the size of the table
    maxColumn = workSheet.max_column
    maxRow = workSheet.max_row

    # Grab the header and store it in the list
    if headerRow == True: 
        for item in range(firstColumn, maxColumn+ 1):
            header.append(workSheet.cell(row = 1, column = item).value)
    
    # get the body and store them in the list of list
    for rowNo in range(firstRow + 1, maxRow +1):
        rowElement = []
        for columnNo in range(1, maxColumn + 1):
            if containHyper == None:
                rowElement.append(workSheet.cell(row = rowNo, column = columnNo).value)
            elif (containHyper == rowNo):
                HyperObject = {}
                target = workSheet.cell(row = rowNo, column = columnNo)
                tag = target.value
                html = target.hyperlink.target
                HyperObject[tag] = html
                rowElement.append(HyperObject)
        print(rowElement)


def main():
    easyHtmlTable('Book 4.xlsx', 'Sheet1', containHyper = 1)

if __name__ == "__main__":
    main()