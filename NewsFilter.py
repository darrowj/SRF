import openpyxl
import HTMLParser
import argparse

'''
This is a simple utility t0 clean up some HTML code from an Excel file.
Usage
  Input -- Input file to process -- inputFileName
        -- Worksheet to process -- worksheetToClean
        -- Column which needs to be cleaned -- columnToClean
        -- File where the results are to be saved -- outputFile
'''
def cleanHtmlValues(inputFileName, worksheetToClean, columnToClean = 0, outputFile = "result.xlsx"):
    from openpyxl import load_workbook
    wb = load_workbook(inputFileName)
    ws = wb[worksheetToClean] # ws is now an IterableWorksheet

    for cellObj in ws.columns[columnToClean]:
        cellValue = HTMLParser.HTMLParser().unescape(cellObj.value)
        #print(cellObj.value)
        cellObj.value = cellValue

    wb.save(outputFile)

def Main():
        parser = argparse.ArgumentParser()
        parser.add_argument("inputFileName", help="Input file to process -- inputFileName.")
        parser.add_argument("worksheetToClean", help="Worksheet to process -- worksheetToClean.")
        parser.add_argument("columnToClean", nargs='?', default=0, help="Column which needs to be cleaned -- columnToClean.", type=int)
        parser.add_argument("outputFile", nargs='?', default="result.xlsx", help="File where the results are to be saved -- outputFile.")

        args = parser.parse_args()

        cleanHtmlValues(args.inputFileName, args.worksheetToClean, args.columnToClean, args.outputFile)

if __name__ == "__main__":
    Main()