from os import path
from meeting_part_manager import MeetingPartManager

def check_files(file1: str = "input_file.xlsx", file2: str = "template.xlsx") -> bool:
    """Return true if both the files are present. Else return False"""

    if path.isfile(file1) and path.isfile(file2):
        return True

    if not path.isfile(file1) and not path.isfile(file2):
        print("Please add a file called '%s' and '%s' in this folder" % (file1, file2))
        print("(Press enter to continue)")
        stdin.read(1)
        return False

    if not path.isfile(file1):
        print("Please add a file called '%s' in this folder" % file1)
        print("(Press enter to continue)")
        stdin.read(1)
        return False

    if not path.isfile("template.xlsx"):
        print("Please add a file called '%s' in this folder" % file2)
        print("(Press enter to continue)")
        stdin.read(1)
        return False


if __name__ == '__main__':
    if not check_files("input_file.xlsx", "template.xlsx"):
        exit()
    MPM = MeetingPartManager("input_file.xlsx", "template.xlsx")
    MPM.shuffle_list(MPM.tuesday_names)
    MPM.shuffle_list(MPM.friday_names)
    MPM.save_to_file()
