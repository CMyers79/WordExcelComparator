#  This program extracts cost and energy savings tables from Word .docx files and searches for user-provided
#  labels that correspond to the cost and energy savings in eProjectBuilder Schedule 4 columns Y and AE.
#  The $/yr and MBTU/yr numbers are then copied to the Comparator.xlsx template and compared to the Schedule 4 values.
#  For building-specific ECM tables, the program searches for the building name in the first row of the table.


import openpyxl
from docx import Document


# helper function returns boolean for string contains int or float
def is_float(string):
    try:
        float(string.replace(",", ""))
        return True
    except ValueError:
        return False


# The docExtractor class should be accessed by the workflow in comparator.py.
class DocExtractor:
    # construct an instance of docExtractor with doc_type 'docx', 'xlsx', or 'epb'
    def __init__(self, doc_type):
        self.data = None
        self.doc_name = None
        self.ecms = [[], []]
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
    def extract(self, ecm_list):
        # load output template for population
        comparator = self.output.get_sheet_by_name('comparator')
        ecm = []
        cost = []
        mbtu = []
        include = []
        baseline = [[], [], [], []]
        savings = [[], [], [], []]
        cost_savings = [[], [], [], []]
        m_v_option = []
        o_m_cost_savings = []
        other_cost_savings = []
        price = []
        spb = []

        # TODO refactor to use class inheritance
        # pull data from columns Y and AE in the epb calculating template
        if self.doc_type == 'epb':
            schedule_4 = self.data.get_sheet_by_name('Sch4-Cost Savings by ECM')
            summary = self.data.get_sheet_by_name('Summary Schedule')

            # populate identifier fields in comparator
            comparator.cell(row=2, column=2, value=summary.cell(row=16, column=3).value)
            # no reliable data source for ESPC Phase
            comparator.cell(row=4, column=2, value=summary.cell(row=9, column=3).value)

            # populate arrays with 250 ECM rows plus the totals
            for i in range(251):
                ecm.append(schedule_4.cell(row=i+8, column=1).value)
                mbtu.append(schedule_4.cell(row=i+8, column=25).value)
                cost.append(schedule_4.cell(row=i+8, column=31).value)
                m_v_option.append(schedule_4.cell(row=i + 8, column=4).value)
                o_m_cost_savings.append(schedule_4.cell(row=i + 8, column=29).value)
                other_cost_savings.append(schedule_4.cell(row=i + 8, column=30).value)
                price.append(schedule_4.cell(row=i + 8, column=32).value)
                spb.append(schedule_4.cell(row=i + 8, column=33).value)
                baseline[0].append(schedule_4.cell(row=i+8, column=5).value)  # kWh
                baseline[1].append(schedule_4.cell(row=i+8, column=6).value)  # kW
                baseline[2].append(schedule_4.cell(row=i+8, column=7).value)  # MBtu
                baseline[3].append(schedule_4.cell(row=i+8, column=10).value)  # kGal
                savings[0].append(schedule_4.cell(row=i + 8, column=15).value)  # kWh
                savings[1].append(schedule_4.cell(row=i + 8, column=17).value)  # kW
                savings[2].append(schedule_4.cell(row=i + 8, column=19).value)  # MBtu
                savings[3].append(schedule_4.cell(row=i + 8, column=27).value)  # kGal
                cost_savings[0].append(schedule_4.cell(row=i + 8, column=16).value)  # kWh
                cost_savings[1].append(schedule_4.cell(row=i + 8, column=18).value)  # kW
                cost_savings[2].append(schedule_4.cell(row=i + 8, column=20).value)  # MBtu
                cost_savings[3].append(schedule_4.cell(row=i + 8, column=28).value)  # kGal

                # add non-blank row indices to include
                if ecm[i] not in (None, ''):
                    include.append(i)

            # remove 'TOTALS' from ecm
            ecm.pop()

            # copy non-blank ECM rows to Comparator.xlsx
            for i in range(len(include) - 1):
                comparator.cell(row=17 + i, column=1, value=ecm[include[i]])
                comparator.cell(row=17 + i, column=2, value=cost[include[i]])
                comparator.cell(row=17 + i, column=7, value=mbtu[include[i]])
                comparator.cell(row=17 + i, column=13, value=price[include[i]])
                comparator.cell(row=17 + i, column=18, value=baseline[0][include[i]])
                comparator.cell(row=17 + i, column=23, value=baseline[1][include[i]])
                comparator.cell(row=17 + i, column=28, value=baseline[2][include[i]])
                comparator.cell(row=17 + i, column=33, value=baseline[3][include[i]])
                comparator.cell(row=17 + i, column=38, value=savings[0][include[i]])
                comparator.cell(row=17 + i, column=43, value=savings[1][include[i]])
                comparator.cell(row=17 + i, column=48, value=savings[2][include[i]])
                comparator.cell(row=17 + i, column=53, value=savings[3][include[i]])
                comparator.cell(row=17 + i, column=58, value=cost_savings[0][include[i]])
                comparator.cell(row=17 + i, column=63, value=cost_savings[1][include[i]])
                comparator.cell(row=17 + i, column=68, value=cost_savings[2][include[i]])
                comparator.cell(row=17 + i, column=73, value=cost_savings[3][include[i]])
                comparator.cell(row=17 + i, column=78, value=m_v_option[include[i]])
                comparator.cell(row=17 + i, column=83, value=o_m_cost_savings[include[i]])
                comparator.cell(row=17 + i, column=88, value=other_cost_savings[include[i]])
                comparator.cell(row=17 + i, column=93, value=spb[include[i]])

            # copy totals to Comparator.xlsx
            comparator.cell(row=15, column=2, value=cost[250])
            comparator.cell(row=15, column=7, value=mbtu[250])
            comparator.cell(row=17 + len(include) - 1, column=2, value=cost[250])
            comparator.cell(row=17 + len(include) - 1, column=7, value=mbtu[250])

            # populate ecm_split list with building names and ecm numbers
            for title in [ecm_title for ecm_title in ecm if ecm_title not in ["", None]]:
                # the ecm number may end in a letter, so look at the penultimate character
                i = len(title) - 2
                # find the last non-decimal character, this is the end of the building name
                while title[i] in "0123456789.":
                    i -= 1
                # add the building name to ecms[0]
                self.ecms[0].append(title[:i + 1].strip().replace('.', ''))

                # find the first digit after the building name, the start of the ecm number
                i = 0
                while title[i] not in "012456789":
                    i += 1
                # add all building ecm numbers to ecms[1]
                self.ecms[1].append(title[i:].strip())

        # pull data from the savings per ECM template
        # elif self.doc_type == 'xlsx':
        #     ecm_savings = self.data.get_sheet_by_name('Sheet1')
        #
        #     # pull data by ECM number starting with B3, copy directly to comparator
        #     row = 3
        #     while ecm_savings.cell(row=row, column=2).value not in (None, ''):
        #         comparator.cell(row=14 + row, column=12, value=ecm_savings.cell(row=row, column=2).value)
        #         comparator.cell(row=14 + row, column=13, value=ecm_savings.cell(row=row, column=9).value)
        #         row += 1
        #
        #     # copy total to Comparator.xlsx
        #     comparator.cell(row=15, column=13, value=ecm_savings.cell(row=row, column=9).value)

        # search for table data within Vol1 or Vol2 Word docx file
        elif self.doc_type == 'docx':
            offset = 0
            if self.doc_name == 'Vol2':
                offset = 1

            # populate list with alias inputs
            alias_cells = [(11 + offset, 2),
                           (11 + offset, 4),
                           (11 + offset, 7),
                           (11 + offset, 9),
                           (11 + offset, 12),
                           (11 + offset, 17),
                           (11 + offset, 22),
                           (11 + offset, 27),
                           (11 + offset, 32),
                           (11 + offset, 37),
                           (11 + offset, 42),
                           (11 + offset, 47),
                           (11 + offset, 52),
                           (11 + offset, 57),
                           (11 + offset, 62),
                           (11 + offset, 67),
                           (11 + offset, 72),
                           (11 + offset, 77),
                           (11 + offset, 82),
                           (11 + offset, 87)]  # value goes in aliases[19]

            aliases = [comparator.cell(row=i, column=j).value for i, j in alias_cells]
            print("Extracting from Vol. 1:")
            print(aliases)
            # iterate through each table in the Word docx (xml) and load the cell values into an array
            for table in self.data.tables:
                cell_text = []
                for row in table.rows:
                    cell_text.append([cell.text for cell in row.cells])

                # search table for aliases from Comparator.xlsx
                for i, row in enumerate(cell_text):
                    for j, cell in enumerate(row):
                        # search each table cell for each alias
                        for alias in aliases:

                            # if building or ECM table found
                            if cell == alias and alias in aliases[2:]:
                                # find the building name in the first row
                                n = 0
                                while n < len(cell_text[0]) - 1 and cell_text[0][n] in ["", None]:
                                    n += 1
                                b_name = cell_text[0][n]
                                b_offset = 0
                                found_offset = False
                                b_ecms = 0

                                # determine the offset of the building in ecm_list and how many ecms at the building
                                for m, name in enumerate(ecm_list[0]):
                                    # must ignore the last word in the building name because it may be 'TC' or 'ECM'
                                    # unless building name is one word
                                    if len(name.split(" ")) > 1:
                                        split_name = " ".join(name.split(" ")[:-1])
                                    else:
                                        split_name = name

                                    if split_name in b_name or b_name in split_name:
                                        # set b_offset to the first row with matching building name
                                        if not found_offset:
                                            b_offset = m
                                            found_offset = True

                                        b_ecms += 1

                            # TODO refactor this repetitive section, standardize comparator table column spacing
                            # if total $ savings is found:
                            if cell == alias and alias == aliases[0]:
                                # pull the numerical value in the cell to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(cell_text[i + 1][j].replace("$", "")):
                                    comparator.cell(row=15,
                                                    column=3 + offset,
                                                    value=float(cell_text[i+1][j].replace(",", "").replace("$", "")))

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1].replace("$", "")):
                                    comparator.cell(row=15,
                                                    column=3 + offset,
                                                    value=float(cell_text[i][j + 1].replace(",", "").replace("$", "")))

                            # if total MBTU is found
                            elif cell == alias and alias == aliases[1]:
                                # pull the numerical value in the cell to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(cell_text[i + 1][j]):
                                    comparator.cell(row=15,
                                                    column=8 + offset,
                                                    value=float(cell_text[i+1][j].replace(",", "")))

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1]):
                                    comparator.cell(row=15,
                                                    column=8 + offset,
                                                    value=float(cell_text[i][j + 1].replace(",", "")))

                            # if ECM $ savings are found
                            elif cell == alias and alias == aliases[2]:
                                # pull the numerical values in the cells to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(cell_text[i + 1][j].replace("$", "")):
                                    k = i

                                    while k < len(cell_text) - 1 and is_float(cell_text[k + 1][j].replace("$", "")) and k - i < b_ecms:
                                        comparator.cell(row=17 + b_offset + k - i,
                                                        column=3 + offset,
                                                        value=float(cell_text[k+1][j].replace(",", "").replace("$", "")))
                                        k += 1

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1].replace("$", "")):
                                    k = j

                                    while k < len(cell_text[i]) - 1 and is_float(cell_text[i][k + 1].replace("$", "")) and k - j < b_ecms:
                                        comparator.cell(row=17 + b_offset + k - j,
                                                        column=3 + offset,
                                                        value=float(cell_text[i][k + 1].replace(",", "").replace("$", "")))
                                        k += 1

                            # if ECM MBTU are found
                            elif cell == alias and alias == aliases[3]:
                                # pull the numerical values in the cells to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(cell_text[i + 1][j]):
                                    k = i

                                    while k < len(cell_text) - 1 and is_float(cell_text[k + 1][j]) and k - i < b_ecms:
                                        comparator.cell(row=17 + b_offset + k - i,
                                                        column=8 + offset,
                                                        value=float(cell_text[k + 1][j].replace(",", "")))
                                        k += 1

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1]):
                                    k = j

                                    while k < len(cell_text[i]) - 1 and is_float(cell_text[i][k + 1]):
                                        comparator.cell(row=17 + b_offset + k - j,
                                                        column=8 + offset,
                                                        value=float(cell_text[i][k + 1].replace(",", "")))

                            # if ECM costs are found
                            elif cell == alias and alias == aliases[4]:
                                # pull the numerical values in the cells to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(cell_text[i + 1][j].replace(",", "").replace("$", "")):
                                    k = i

                                    while k < len(cell_text) - 1 and is_float(
                                            cell_text[k + 1][j].replace(",", "").replace("$", "")) and k - i < b_ecms:
                                        comparator.cell(row=17 + b_offset + k - i,
                                                        column=14 + offset,
                                                        value=float(cell_text[k + 1][j].replace(",", "").replace("$", "")))
                                        k += 1

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1]):
                                    k = j

                                    while k < len(cell_text[i]) - 1 and is_float(cell_text[i][k + 1]):
                                        comparator.cell(row=17 + b_offset + k - j,
                                                        column=14 + offset,
                                                        value=float(cell_text[i][k + 1].replace(",", "").replace("$", "")))

                            # if baseline kWh is found
                            elif cell == alias and alias == aliases[5]:
                                # pull the numerical values in the cells to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(cell_text[i + 1][j].replace(",", "")):
                                    k = i

                                    while k < len(cell_text) - 1 and is_float(cell_text[k + 1][j].replace(",", "")) and k - i < b_ecms:
                                        comparator.cell(row=17 + b_offset + k - i,
                                                        column=19 + offset,
                                                        value=float(cell_text[k + 1][j].replace(",", "")))
                                        k += 1

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1]):
                                    k = j

                                    while k < len(cell_text[i]) - 1 and is_float(cell_text[i][k + 1]):
                                        comparator.cell(row=17 + b_offset + k - j,
                                                        column=19 + offset,
                                                        value=float(cell_text[i][k + 1].replace(",", "")))

                            # if baseline kW is found
                            elif cell == alias and alias == aliases[6]:
                                # pull the numerical values in the cells to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(
                                        cell_text[i + 1][j].replace(",", "")):
                                    k = i

                                    while k < len(cell_text) - 1 and is_float(cell_text[k + 1][j].replace(",", "")) and k - i < b_ecms:
                                        comparator.cell(row=17 + b_offset + k - i,
                                                        column=24 + offset,
                                                        value=float(cell_text[k + 1][j].replace(",", "")))
                                        k += 1

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1]):
                                    k = j

                                    while k < len(cell_text[i]) - 1 and is_float(cell_text[i][k + 1]):
                                        comparator.cell(row=17 + b_offset + k - j,
                                                        column=24 + offset,
                                                        value=float(cell_text[i][k + 1].replace(",", "")))

                            # if baseline gas is found
                            elif cell == alias and alias == aliases[7]:
                                # pull the numerical values in the cells to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(cell_text[i + 1][j].replace(",", "")):
                                    k = i

                                    while k < len(cell_text) - 1 and is_float(cell_text[k + 1][j].replace(",", "")) and k - i < b_ecms:
                                        comparator.cell(row=17 + b_offset + k - i,
                                                        column=29 + offset,
                                                        value=float(cell_text[k + 1][j].replace(",", "")))
                                        k += 1

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1]):
                                    k = j

                                    while k < len(cell_text[i]) - 1 and is_float(cell_text[i][k + 1]):
                                        comparator.cell(row=17 + b_offset + k - j,
                                                        column=29 + offset,
                                                        value=float(cell_text[i][k + 1].replace(",", "")))

                            # if baseline kgal is found
                            elif cell == alias and alias == aliases[8]:
                                # pull the numerical values in the cells to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(cell_text[i + 1][j].replace(",", "")):
                                    k = i

                                    while k < len(cell_text) - 1 and is_float(
                                            cell_text[k + 1][j].replace(",", "")) and k - i < b_ecms:
                                        comparator.cell(row=17 + b_offset + k - i,
                                                        column=34 + offset,
                                                        value=float(cell_text[k + 1][j].replace(",", "")))
                                        k += 1

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1]):
                                    k = j

                                    while k < len(cell_text[i]) - 1 and is_float(cell_text[i][k + 1]):
                                        comparator.cell(row=17 + b_offset + k - j,
                                                        column=34 + offset,
                                                        value=float(cell_text[i][k + 1].replace(",", "")))

                            # if kWh savings is found
                            elif cell == alias and alias == aliases[9]:
                                # pull the numerical values in the cells to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(
                                        cell_text[i + 1][j].replace(",", "").replace("$", "")):
                                    k = i

                                    while k < len(cell_text) - 1 and is_float(cell_text[k + 1][j].replace(",", "")) and k - i < b_ecms:
                                        comparator.cell(row=17 + b_offset + k - i,
                                                        column=39 + offset,
                                                        value=float(cell_text[k + 1][j].replace(",", "").replace("$", "")))
                                        k += 1

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1]):
                                    k = j

                                    while k < len(cell_text[i]) - 1 and is_float(cell_text[i][k + 1]):
                                        comparator.cell(row=17 + b_offset + k - j,
                                                        column=39 + offset,
                                                        value=float(cell_text[i][k + 1].replace(",", "").replace("$", "")))

                            # if kW savings is found
                            elif cell == alias and alias == aliases[10]:
                                # pull the numerical values in the cells to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(cell_text[i + 1][j].replace(",", "").replace("$", "")):
                                    k = i

                                    while k < len(cell_text) - 1 and is_float(
                                            cell_text[k + 1][j].replace(",", "")) and k - i < b_ecms:
                                        comparator.cell(row=17 + b_offset + k - i,
                                                        column=44 + offset,
                                                        value=float(cell_text[k + 1][j].replace(",", "").replace("$", "")))
                                        k += 1

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1]):
                                    k = j

                                    while k < len(cell_text[i]) - 1 and is_float(cell_text[i][k + 1]):
                                        comparator.cell(row=17 + b_offset + k - j,
                                                        column=44 + offset,
                                                        value=float(cell_text[i][k + 1].replace(",", "").replace("$", "")))

                            # if gas savings is found
                            elif cell == alias and alias == aliases[11]:
                                # pull the numerical values in the cells to either the right or bottom of the alias
                                if i < len(cell_text) - 1 and is_float(cell_text[i + 1][j].replace(",", "").replace("$", "")):
                                    k = i

                                    while k < len(cell_text) - 1 and is_float(
                                            cell_text[k + 1][j].replace(",", "")) and k - i < b_ecms:
                                        comparator.cell(row=17 + b_offset + k - i,
                                                        column=49 + offset,
                                                        value=float(cell_text[k + 1][j].replace(",", "").replace("$", "")))
                                        k += 1

                                elif j < len(cell_text[i]) - 1 and is_float(cell_text[i][j + 1]):
                                    k = j

                                    while k < len(cell_text[i]) - 1 and is_float(cell_text[i][k + 1]):
                                        comparator.cell(row=17 + b_offset + k - j,
                                                        column=49 + offset,
                                                        value=float(cell_text[i][k + 1].replace(",", "").replace("$", "")))

                            # if kgal savings is found
                            elif cell == alias and alias == aliases[12]:
                                # pull the numerical values in the cells to either the right or bottom of the alias
                                m = i
                                n = j
                                # skip blank cells and check the first populated cell
                                while cell_text[m + 1][j] == '\xa0' and m < len(cell_text) - 1:
                                    m += 1
                                while cell_text[i][n + 1] == '\xa0' and n < len(cell_text[i]) - 1:
                                    n += 1

                                if is_float(cell_text[m + 1][j].replace(",", "").replace("$", "")):
                                    k = i

                                    while k < len(cell_text) - 1 and (is_float(cell_text[k + 1][j].replace(",", "")) or cell_text[k + 1][j] == '\xa0') \
                                            and k - i < b_ecms:
                                        if cell_text[k + 1][j] == '\xa0':
                                            k += 1
                                            continue
                                        else:
                                            comparator.cell(row=17 + b_offset + k - i,
                                                            column=54 + offset,
                                                            value=float(cell_text[k + 1][j].replace(",", "").replace("$", "")))
                                        k += 1

                                elif is_float(cell_text[i][n + 1].replace(",", "").replace("$", "")):
                                    k = j

                                    while k < len(cell_text[i]) - 1 and (is_float(cell_text[i][k + 1].replace(",", "")) or cell_text[i][k + 1] == '\xa0') \
                                            and k - j < b_ecms:
                                        if cell_text[i][k + 1] == '\xa0':
                                            k += 1
                                            continue
                                        else:
                                            comparator.cell(row=17 + b_offset + k - j,
                                                            column=54 + offset,
                                                            value=float(cell_text[i][k + 1].replace(",", "").replace("$", "")))

        self.output.save('Comparator.xlsx')
