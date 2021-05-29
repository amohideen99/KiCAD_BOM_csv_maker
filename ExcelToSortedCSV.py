import xlrd
import sys
from collections import defaultdict

'''
    Directions for Use:
        1. go to KiCad symbol fields and uncheck all "group by..."
        2. select all and copy into Excel file
        3. save Excel file as .xls NOT .xlsx
        4. go to directory with python script (ie "My PC\Documents\BOM Creation Tools")
        5. run python file with
            argument 1 as the directory for .xls file from KiCad
            argument 2 as output directory for CSV file
            argument 3 as multiplier for number of components (if you need to make multiple boards otherwise defaults to 1)

    Expected Excel File Format
        Column A - Schematic Reference
        Column B - Component Value
        Column E - Digikey Part #
        Column F - Mouser Part #

    CSV Output File Format
        value, matching references, digikey #, mouser #, quantity per board, total order quantity

'''
# open excel file
wb = xlrd.open_workbook(sys.argv[1])
sheet = wb.sheet_by_index(0)

# instantiate dicts for holding value, reference, digikey #, and mouser #
values_dict = defaultdict(list)
digikey_dict = {}
mouser_dict = {}

# go thru excel file row by row and fill up dictionaries
for i in range(sheet.nrows):
    values_dict[sheet.cell_value(i,1)].append(sheet.cell_value(i, 0).strip())
    digikey_dict[sheet.cell_value(i,1)] = sheet.cell_value(i, 4).strip()
    mouser_dict[sheet.cell_value(i,1)] = str(sheet.cell_value(i, 5)).strip()

# method for returning the list of references
# as "C1 C2 C3" etc.
def listToString(s):
    str1 = ""
    # traverse in the string
    for ele in s:
        str1 += ele + " "
    # return string
    return str1[:len(str1)-1]

# logic for multiplier
# if a multiplier provided, use it
# otherwise default to 1
if not (len(sys.argv) == 3):
    multiplier = int(sys.argv[3]);
else:
    multiplier = 1;

# go to output path and create/clear text file
# open the file for writing
file_path = sys.argv[2] + "\BOM_sorted_CSV.txt"
file = open(file_path, "w").close()
file = open(file_path, "a+")

# go through the dictionaries and place appropriate values in file
for key in values_dict:
    out = str(key) + "," + listToString(values_dict[key]) + ","
            + str(digikey_dict[key]) + "," + str(mouser_dict[key]) + ","
            + str(len(values_dict[key])) + ","
            + str(len(values_dict[key]) * multiplier) + "\n"

    file.write(out)
file.close()
