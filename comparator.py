#  This program extracts cost and energy savings tables from Word .docx files and searches for user-provided
#  labels that correspond to the cost and energy savings in eProjectBuilder Schedule 4 columns Y and AE.
#  The $/yr and MBTU/yr numbers are then copied to the Comparator.xlsx template and compared to the Schedule 4 values.

from docExtract import docExtractor
import os

if __name__ == "__main__":

    directory = os.getcwd()
    epb_workflow = ecm_savings_workflow = vol_1_workflow = vol_2_workflow = None

    if os.path.isfile(directory + 'epb.xlsx'):
        epb_workflow = docExtractor('epb')

    if os.path.isfile(directory + 'ECMSavings.xlsx'):
        ecm_savings_workflow = docExtractor('xlsx')

    if os.path.isfile(directory + 'Vol1.docx'):
        vol_1_workflow = docExtractor('docx')

    if os.path.isfile(directory + 'Vol2.docx'):
        vol_2_workflow = docExtractor('docx')


