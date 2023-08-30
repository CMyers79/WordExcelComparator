#  This program extracts cost and energy savings tables from Word .docx files and searches for user-provided
#  labels that correspond to the cost and energy savings in eProjectBuilder Schedule 4 columns Y and AE.
#  The $/yr and MBTU/yr numbers are then copied to the Comparator.xlsx template and compared to the Schedule 4 values.

import openpyxl
from docx import Document


# helper function to check if a string contains an int or float
def is_float(string):
    try:
        float(string.replace(",", ""))
        return True
    except ValueError:
        return False


# The docExtractor class should be accessed by the workflow in comparator.py.
class docExtractor:
    # construct an instance of docExtractor with doc_type 'docx', 'xlsx', or 'epb'
    def __init__(self, doc_type):
        self.data = None
        self.doc_name = None
        self.output = openpyxl.load_workbook('Comparator.xlsx')
        self.doc_type = doc_type

    # load file at filepath passed in
    def load(self, filepath):
        # load file into self.data
        if self.doc_type == 'epb' or self.doc_type == 'xlsx':
            self.data = openpyxl.load_workbook(filepath, data_only=True)
        elif self.doc_type == 'docx':
            self.data = Document(filepath)
            if filepath == 'Vol1.docx':
                self.doc_name = 'Vol1'
            elif filepath == 'Vol2.docx':
                self.doc_name = 'Vol2'

    # extract table data from self.data by self.doc_type
    def extract(self):
        # load output template for population
        comparator = self.output.get_sheet_by_name('comparator')
        ecm = []
        cost = []
        mbtu = []
        include = []

        # pull data from columns Y and AE in the epb calculating template
        if self.doc_type == 'epb':
            schedule_4 = self.data.get_sheet_by_name('Sch4-Cost Savings by ECM')

            # populate arrays with 250 ECM rows plus the total
            for i in range(251):
                ecm.append(schedule_4.cell(row=i+8, column=1).value)
                mbtu.append(schedule_4.cell(row=i+8, column=25).value)
                cost.append(schedule_4.cell(row=i+8, column=31).value)
                # add non-blank row indices to include
                if ecm[i] not in (None, ''):
                    include.append(i)
            # copy non-blank ECM rows to Comparator.xlsx
            for i in range(len(include) - 1):
                comparator.cell(row=9 + i, column=1, value=ecm[include[i]])
                comparator.cell(row=9 + i, column=2, value=cost[include[i]])
                comparator.cell(row=9 + i, column=7, value=mbtu[include[i]])

            # copy totals to Comparator.xlsx
            comparator.cell(row=7, column=2, value=cost[250])
            comparator.cell(row=7, column=7, value=mbtu[250])

        # pull data from the savings per ECM template
        elif self.doc_type == 'xlsx':
            ecm_savings = self.data.get_sheet_by_name('Sheet1')

            # pull data by ECM number starting with B3, copy directly to comparator
            row = 3
            while ecm_savings.cell(row=row, column=2).value not in (None, ''):
                comparator.cell(row=6 + row, column=12, value=ecm_savings.cell(row=row, column=2).value)
                comparator.cell(row=6 + row, column=13, value=ecm_savings.cell(row=row, column=9).value)
                row += 1

            # copy total to Comparator.xlsx
            comparator.cell(row=7, column=13, value=ecm_savings.cell(row=row, column=9).value)

        # search for table data within Vol1 or Vol2 Word docx file
        elif self.doc_type == 'docx':
            offset = 0
            if self.doc_name == 'Vol2':
                offset = 1

            # iterate through each table in the Word docx (xml) and load the cell values into an array
            for table in self.data.tables:
                cell_text = []
                for row in table.rows:
                    cell_text.append([cell.text for cell in row.cells])

                # search table for aliases from Comparator.xlsx
                for i, row in enumerate(cell_text):
                    for j, cell in enumerate(row):
                        # search for
                        if cell == comparator.cell(row=3 + offset, column=4).value:
                            # find the numerical value in the cell to either the right or bottom of the alias
                            if i < len(cell_text) - 1 and is_float(cell_text[i + 1][j]):
                                comparator.cell(row=7, column=8 + offset, value=float(cell_text[i+1][j].replace(",", "")))
                            elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1]):
                                comparator.cell(row=7, column=8 + offset, value=float(cell_text[i][j + 1].replace(",", "")))




        self.output.save('Comparator.xlsx')
