"""Contains Class Meeting Part"""
import datetime
from openpyxl import Workbook, load_workbook


class MeetingPart:
    """Holds basic data for a single meeting part,
    and functions to read and write to excel file

    Variables:
        part_name (str): Name of the Meeting Part
        input_sheet (openpyxl.load_workbook.worksheet): Sheet from the 'input_file' excel document
        column (chr): Column from the input file that contains the relevant data
        tuesday (bool): whether the part is on a Tuesday or not (Friday)
        start_date (datetime.date): the start date for part assignment
        names (list): list of all the names of the people that can be assigned this part
        shuffled_names (list): same list after being randomized

    Methods:
        read_names() : read the names of people from the input excel sheet and
            store them in the 'names' list
        set_shuffled_names(shuffled_names): sets the value for the shuffled names list
        write_to_file(row1, row2, output_sheet, output_sheet2, start_date):
            Writes

    """


    def __init__(self, input_sheet, column: chr, tuesday: bool, start_date: datetime.date) -> MeetingPart:
        """Initialize the Meeting Part Class object with apropriate variables"""

        self.input_sheet = input_sheet
        self.column = column
        self.tuesday = tuesday
        self.start_date = start_date
        self.names = []
        self.shuffled_names = []
        self.read_names()


    def read_names(self):
        """Read the names from the input sheet and store in list names"""

        for cell in self.in_sheet[self.column]:
            if cell.value is not None:
                self.names.append(cell.value)
        self.part_name = names[0]
        self.names.pop(0)


    def set_shuffled_names(self, shuffled_names: list):
        """Set the value of the list shuffled_names"""

        self.shuffled_names = shuffled_names


    def write_to_file (self, row1: int, row2: int, output_sheet, output_sheet2, start_column: int):
        """Write data to output file"""

        if tuesday:
            for i in range(len(self.shuffled_names)):
                output_sheet.cell(row=row1, column=start_column+i*2).value = self.shuffled_names[i]
                output_sheet2.cell(row=row2+i, column=2).value = self.shuffled_names[i]
        else:
            for i in range(len(self.shuffled_names)):
                output_sheet.cell(row=row1, column=4+i*2).value = wt_readers[i]
                output_sheet2.cell(row=row2+i, column=1).value = wt_readers[i]




if __name__ == '__main__':

    input_wb = load_workbook("input_file.xlsx")
    output_wb = load_workbook("template.xlsx")
    input_sheet = input_wb.worksheets[0]
    output_sheet = output_wb.worksheets[0]
    output_sheet2 = output_wb.worksheets[1]
    wt_readers = []
    cbs_readers = []

    chairman = MeetingPart("Chairman", input_sheet, 'A', True, datetime.date.today())
    chairman.names
