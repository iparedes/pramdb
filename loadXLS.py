from openpyxl import load_workbook
from openpyxl.utils.cell import *
import re


class AssessmentXLS:

    def __init__(self):
        self.book = load_workbook('assessment.xlsx')
        self.ws = self.book['Assessment']


        r=self.book.defined_names['AssetName']
        dests=list(r.destinations)[0][1]
        self.assetName=self.ws[dests].value

        r=self.book.defined_names['AssetType']
        dests=list(r.destinations)[0][1]
        self.assetType=self.ws[dests].value

        self.impacts = self.table_to_list('TblImpact')
        self.asls = self.table_to_list('TblAssessment')
        self.scenarios=self.table_to_list('TblScenarios')

    # Needs a table identifier
    # Returns a list of where each element corresponds to a row in the list, being a dictionary with the title of the columns as keys
    # if the first column is blank, skips the row
    def table_to_list(self, tableId):

        tab = self.ws.tables[tableId]
        titles = tab.column_names
        # range=self.book.defined_names['TblImpact'].value
        rango = self.ws.tables[tableId].ref
        pattern = '([A-Z]+)([0-9]+)\:([A-Z]+)([0-9]+)'
        res = re.match(pattern, rango)

        a = res.group(1)
        colIni = column_index_from_string(a)

        rowIni = int(res.group(2))
        rowIni += 1  # skip the title

        a = res.group(3)
        colEnd = column_index_from_string(a)

        rowEnd = int(res.group(4))

        records = []
        for row in range(rowIni, rowEnd + 1):
            item = {}
            cont = 0
            for t in titles:
                item[t] = self.ws.cell(row, colIni + cont).value
                cont += 1
            if item[titles[0]]:
                records.append(item)
        return records

    def range_char(self, start, stop):
        return (chr(n) for n in range(ord(start), ord(stop) + 1))


    # Creates a sheet with a given name
    # Overwrites the sheet if it already exists
    # sets the new sheet to the active sheet
    def create_sheet(self,name):
        snames=self.book.sheetnames
        for s in snames:
            if s==name:
                sheet=self.book.get_sheet_by_name(name)
                self.book.remove_sheet(sheet)
        new_sheet=self.book.create_sheet(name)
        self.book.save("assessment.xlsx")

    def set_cell(self,row,column,value,sheetname=""):
        if sheetname:
            ws=self.book.get_sheet_by_name(sheetname)
        else:
            ws=self.book.active
        ws.cell(row,column).value=value

    def add_vector(self,row,column,vector,sheetname=""):
        if sheetname:
            ws=self.book.get_sheet_by_name(sheetname)
        else:
            ws=self.book.active
        col=column
        for v in vector:
            ws.cell(row=row,column=col).value=v
            col+=1
        #self.book.save("assessment.xlsx")





