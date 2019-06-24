"""Contains Meeting Part Manager class"""
import datetime
import itertools
from os import path
from time import time
from random import shuffle
from inflect import engine
from openpyxl import Workbook
from openpyxl import load_workbook
from meeting_part import MeetingPart


class MeetingPartManager:
    """Holds multiple objects of the MeetingPart class.
    """

    def __init__(self, in_file_name:str = "input_file.xlsx", out_file_name:str = "template.xlsx") -> 'MeetingPartManager':
        """Initialize the Meeting Part Manager class with input file name, output
        file name, and the cell that contains the start date in the input file"""

        input_wb = load_workbook(in_file_name)
        self.input_sheet = input_wb.worksheets[0]
        self.output_wb = load_workbook(out_file_name)
        self.output_sheet = self.output_wb.worksheets[0]
        self.output_sheet2 = self.output_wb.worksheets[1]
        # self.date_cell = date_cell
        self.index_tuesday = {}
        self.index_friday  = {}
        self.meeting_parts = []
        self.tuesday_names = []
        self.friday_names  = []

        self.get_start_date()
        self.create_parts()
        # self.shuffle_list(self.tuesday_names)


    def get_start_date(self):
        """Read the start date from the input file. Default to todays date otherwise."""

        self.no_of_parts = 9
        self.start_date = datetime.date.today()
        found_date = False
        second_row = self.input_sheet[2]
        second_row = [second_row[i].value for i in range(len(second_row))]      # Convert row cell to list
        for i in range(len(second_row)):
            if type(second_row[i]) == datetime.datetime:
                self.start_date = second_row[i]
                self.no_of_parts = i - 1            # Do not include the labels column 'A' (-1)
                found_date = True
                break

        if not found_date:
            print("Defaulting to todays date as input for the start date...")


    def create_parts(self):
        """Create objects of class MeetingPart and store them in the meeting_parts list.
        Also add all the names to the names list"""

        for i in range(1, self.no_of_parts):
            col = chr(ord('A') + i)
            mp = MeetingPart(self.input_sheet, col)
            self.meeting_parts.append(mp)
            if mp.tuesday:
                self.tuesday_names.append(mp.get_names())
                self.index_tuesday[len(self.tuesday_names)-1] = i
            else:
                self.friday_names.append(mp.get_names())
                self.index_friday[len(self.friday_names)-1] = i



    def shuffle_list(self, names: list):
        maxTimeLimit = 0.05
        timeExceded = True
        while timeExceded:
            timeExceded = False
            startTime = time()
            shuffle(names[0])
            for i in range(len(names)):
                transposedNames = [list(x) for x in itertools.zip_longest(*names[:i+1])]
                while self.checkDupCols(transposedNames):
                    shuffle(names[i])
                    transposedNames = [list(x) for x in itertools.zip_longest(*names[:i+1])]
                    if( time()-startTime > maxTimeLimit):
                        print("Time taken for iteration", _, "was  = ", time()-startTime,)
                        print("Restarting iteration")
                        timeExceded = True
                        break
                if timeExceded:
                    break


    def checkDupCols(self, arr):          #Actually checks duplicate rows, but for a transposed array
        for row in arr:
            seen = set()
            for x in row:
                if x in seen:
                    return True
                if x is not None:
                    seen.add(x)
        return False


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
    MPM = MeetingPartManager("input_file.xlsx", "template.xlsx")
    MPM.no_of_parts
    MPM.start_date
    MPM.tuesday_names
    MPM.friday_names
    MPM.shuffle_list(MPM.tuesday_names)
    MPM.tuesday_names
    MPM.start_date
    # MPM.save_to_file()
