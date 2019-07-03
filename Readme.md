# Automatic Scheduler
A program that creates a schedule for multiple meeting parts.

---

## Instructions
To use the program simply fill in the details on the input file, and run the `.exe`. 
The program will :
1. Read the list of names for each part on the input file.
2. Randomly shuffle the list
3. Ensure that no one is scheduled to do 2 parts on the same day
4. Make a copy of the template file
5. Write the shuffeled list of names into the file
6. Save it to a seperate output file.

> The files `input_file.xlsx` and `template.xlsx` must be present in the same folder as the `.exe` file for the program to work.

### What and What not to change
Anything that is aesthetics related can be changed, like the background color, font, font size, cell size, etc. This applies to both the input file and template file. Most text can also be altered without any issues.

Only thing related to the positions of certain cells should not be changed. The program makes certain assumptions like: 
* The first input will be read from column B of the input file.
* The start date will be present somewhere in row 2 of the input file.
* The template file will contain atleast 2 sheets.
* The first output will be written to column C of the output file.
* The dates will be written to row 2 of the template file. 

While the program will still run for the most part if these things are changed, but it might produce unexpected results, so avoid inserting, deleting or modifying the first few rows/columns for the file. That said, if a change is necessary, do try and make the change and see if the end result is as expected.

## Input File
All the details used by the program are read off of a file called `input_file.xlsx`. The file needs to have excatly the same name for it to work. 
By default, the excel file has 9 columns of individual meeting parts. Feel free to add more columns or delete a few columns. But only columns before the 'date' field will be read. 

The first 3 rows of each column must be filled out. Any missing data in the first 3 rows will cause that entire column to be ignored. They may be intentionally left blank to skip that column.

Part name | Chairman Tuesday| Chairman Friday | Treasures 1 | Treasures 2 | Living as C 1 | WT Conductor | CBS Conductor | WT Readers | CBS Readers | Enter Start date
-|-|-|-|-|-|-|-|-|-|-|
Tuesday/Friday | Tuesday | Friday | Tuesday | Tuesday | Tuesday | Friday | Tuesday | Friday | Tuesday | 11/07/2019
Row Number | 3 | 3 | 4 | 5 | 10 | 12 | 12 | 13 | 13

##### Part Name (row 1)
Enter the name/title of each meeting part. This is purely for the users refference, and as such can be edited as desired. But ensure that there is atlesat some text entered or the column won't be read.

##### Tuesday/Friday (row 2)
Specify if the meeting part is to be scheduled on a tuesday or a friday. If either **Tuesday**/**tuesday** is entered it will be scheduled to a tuesday. If **anything else** is entered in this field - including names of any other days of the week - it will default to friday.

##### Row Number (row 3) 
Tell the program which row number in `template.xlsx` file this meeting part corresponds to . For example, 'Treasures 1' is on row 4 in the template file. So enter '4' as the row number here. Similarly 'WT Reader' is on row 13, so enter row number 13 for the WT Reader part.

##### Names (row 4 onwards)
From row 4 onwards, you can start entering the list of names that you want to schedule for the given part. This list can be as long as you like (even if it goes beyond the colored cells). 
Names can be repeated if needed. **eg:-** for WT conductor GJ is repeated many times. Also if any list is too small the whole list can be doubled or trippled to try and make its length closer to some of the longer parts on the list (purely optional). 
***Warning***
The program atempts to always create a schedule without conflicts, meaning that no person will be scheduled to do 2 parts on the same day. So it is possible to end up with a list of names that is impossible to schedule without conflict. **eg:-** If the WT conductor is filled with GJ, and GJ is also added to the list of Friday Chairman, then no possible combination can be made without a clashing the chairman and WT conductor. So the program will simply quit without producing any output. 

##### Start Date (last column)
Enter the first date of the schedule. It can be any day. The program will automatically find the next closest friday/tuesday from the given date. This date must be entered on row 2, the column number does not matter.
The program simply uses this date to add dates to the top of the output file. 
If the date field is left blank, the start date will default to today's date. 
***Warning***
Because of the way the program was written, it may be necessary to fill in this start date instead of leaving it blank. If the program does not work try adding a start date and see if it fixes the error.


## Template File
The progam creates a copy of `template.xlsx`, fills it with the randomized schedule, and saves it to a file called `output.xlsx`. The only purose of the template file is to have a output that can be used with almost no changes required (Some things like deleting excess columns will have to be done manually).
It is possible to have a completely template file and the programn will run perfectly file. The only requirement is that **the file exists** and the file has **2 sheets** (the program will not run if the template file has only a single sheet)

##### Sheet 1
Sheet 1 will be where the names are writen assuming that the regular meeting schedule template is used. Certain assumptions made are:
1. There are alternating Tuesday and Friday columns
2. It starts on a Tuesday
3. It starts from Column C
4. Row 2 is for the date.

Based on these assumptions the program will only write output from column C onwards, and it will write the date on row 2. 
Each part will be placed on only even columns or odd columns depending on wether the part was specified as being on tuesday or friday. And each part will be written to the row that it corresponds to based on the 'row number' entered in the input file.

##### Sheet 2 
Here the same list that is written to the schedule on sheet 1 is written again, but in a much more simple format. This repeated list is provided if you want to make your own schedule without, or for refference for future months. 
It follows the same format as the input file, where the first row will contain the name of the meeting part, and the list of names in the new randomized order below it. 
This entire sheet is left blank, and will be filled in by the program.

## Output File
The output file contains the final output of the program. The names, taken from input, will be shuffled randomly, while ensuring that there are no clashes. The result is a randomized schedule in which no person is assigned 2 parts on the same day.
The file has 2 sheets. Sheet 1 contains the schedule as is normally created and shared. This can be used as is, or slightly modified for aesthetics, or only the data taken and coppied into a different template file. Sheet 2 contains the same randomized list in a much simpler format, incase you want to build your own schedule, or incase you want to keep it as refference for future months.
