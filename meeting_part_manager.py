"""Contains Meeting Part Manager class"""
import datetime
from os import path
from inflect import engine
from openpyxl import Workbook, load_workbook


class MeetingPartManager:
    """Holds multiple objects of the MeetingPart class.
    """

    def __init__(self, in_file_name:str = "input_file.xlsx", out_file_name:str
                = "template.xlsx", date_cell:str = "J2") -> 'MeetingPartManager':
        """Initialize the Meeting Part Manager class with input file name, output
        file name, and the cell that contains the start date in the input file"""

        self.in_file_name = in_file_name
        self.out_file_name = out_file_name  # Misnommer. Should have been template_file_name. But it's ok, don't change.
        self.date_cell = date_cell
        self.input_wb = load_workbook(in_file_name, read_only=True)
        self.output_wb = load_workbook(out_file_name, read_only=False)
        self.input_sheet = self.input_wb.worksheets[0]
        self.output_sheet = self.output_wb.worksheets[0]
        self.output_sheet2 = self.output_wb.worksheets[1]
        self.get_start_date()


    def get_start_date(self):
        """Read the start date from the input file. Default to todays date otherwise."""

        entered_date = self.input_sheet[self.date_cell].value    # Local Variable
        if (entered_date is not None) and not (type(entered_date) == datetime.datetime):
            print("Unable to read the entered date.")
            print("In the input file, please change the format of the cell for the start date to type 'date' (short date or long date will do)\n")

        if (entered_date is not None) and (type(entered_date) == datetime.datetime):
            start_date = entered_date                       # Local variable
        else:
            print("Defaulting to todays date as input for the start date...")
            start_date = datetime.date.today()
        self.start_date = start_date


    def save_to_file(self):
        # Get the next tuesday and friday from the given date (inclusive)
        tuesday = self.start_date + datetime.timedelta((1 - self.start_date.weekday()) % 7)
        friday = self.start_date + datetime.timedelta((4 - self.start_date.weekday()) % 7)

        tuesday
        friday

        # Check wether tuesday comes first or not
        if (tuesday < friday):
            tuesday_first = True
            start_column = 3
        else:
            tuesday_first = False
            start_column = 5


        ###################################################
        # Add code to save names from each part to file.  #
        # Call the write_to_file() funciton for each part #
        ###################################################


        #Write dates to the top of the file.
        # max_weeks = max(len(wt_readers), len(cbs_readers))
        max_weeks = 15
        max_weeks
        p = engine()
        for i in range(max_weeks):
            date1 = p.ordinal(tuesday.day) + " " + tuesday.strftime('%B')
            date2 = p.ordinal(friday.day) + " " + friday.strftime('%B')
            self.output_sheet.cell(row=2, column=start_column+(i*2)).value = date1
            self.output_sheet.cell(row=2, column=4+(i*2)).value = date2

            tuesday += datetime.timedelta(7)
            friday += datetime.timedelta(7)

        # Save the whole thing to a file named output with renaming if needed
        output_file_path = "Output"
        index = ''
        while path.isfile(output_file_path+index+".xlsx"):
            if index:
                index = '(' + str(int(index[1:-1]) + 1) + ')'
            else:
                index = '(1)'

        self.output_wb.save(output_file_path+index+".xlsx")
        print("Output file created successfully:", output_file_path+index+".xlsx")
        print("(Press enter to close)")
        # stdin.read(1)



if __name__ == '__main__':
    MPM = MeetingPartManager("input_file.xlsx", "template.xlsx", "J2")
    MPM.start_date
    MPM.save_to_file()
