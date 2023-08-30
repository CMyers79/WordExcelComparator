#  This program extracts cost and energy savings tables from Word .docx files and searches for user-provided
#  labels that correspond to the cost and energy savings in eProjectBuilder Schedule 4 columns Y and AE.
#  The $/yr and MBTU/yr numbers are then copied to the Comparator.xlsx template and compared to the Schedule 4 values.

import openpyxl
from docx import Document


# The docExtractor class should be accessed by the workflow in comparator.py.
class docExtractor:
    # construct an instance of docExtractor with doc_type 'docx', 'xlsx', or 'epb'
    def __init__(self, doc_type):
        self.data = None
        self.output = openpyxl.load_workbook('Comparator.xlsx')
        self.doc_type = doc_type

    # load file at filepath passed in
    def load(self, filepath):
        # load file into self.data
        if self.doc_type == 'epb' or self.doc_type == 'xlsx':
            self.data = openpyxl.load_workbook(filepath, data_only=True)
        elif self.doc_type == 'docx':
            self.data = Document(filepath)

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
                comparator.cell(row=10 + i, column=1, value=ecm[include[i]])
                comparator.cell(row=10 + i, column=2, value=cost[include[i]])
                comparator.cell(row=10 + i, column=6, value=mbtu[include[i]])

            # copy totals to Comparator.xlsx
            comparator.cell(row=8, column=2, value=cost[250])
            comparator.cell(row=8, column=6, value=mbtu[250])

        # pull data from the savings per ECM template
        elif self.doc_type == 'xlsx':
            ecm_savings = self.data.get_sheet_by_name('Sheet1')

            # pull data by ECM number starting with B3, copy directly to comparator
            row = 3
            while ecm_savings.cell(row=row, column=2).value not in (None, ''):
                comparator.cell(row=7 + row, column=10, value=ecm_savings.cell(row=row, column=2).value)
                comparator.cell(row=7 + row, column=11, value=ecm_savings.cell(row=row, column=9).value)
                row += 1

            # copy total to Comparator.xlsx
            comparator.cell(row=8, column=11, value=ecm_savings.cell(row=row, column=9).value)

        # search for table data within Vol1 or Vol2 Word docx file
        elif self.doc_type == 'docx':
            pass

        self.output.save('Comparator.xlsx')
