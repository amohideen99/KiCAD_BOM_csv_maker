# KiCAD BOM .csv Maker
BOM creation tool for KiCAD projects. Makes .csv file for direct upload to DigiKey, Mouser, etc.

## Directions for Use

### 1. Copy Symbol Tables from KiCAD

>Be sure to uncheck all "group by" options

### 2. Paste into Excel File

### 3. Save Excel File as .xls **_NOT .xlsx_**

### 4. Run python script from command line
```
    argument 1 - location for .xls file
    argument 2 - output directory for CSV file
    argument 3 - multiplier for number of boards (defaults to 1)
```

### Input .xls File Expected Format

Program expects following column specifications:
```
    Column A - Schematic Reference
    Column B - Component Value
    Column E - Digikey Part #
    Column F - Mouser Part #
```
### Output .csv File Format

Program formats the output file as follows:
```
  value, matching references, digikey #, mouser #, quantity per board, total order quantity
```
