# excel_to_txt.py
""": The program should open a spreadsheet and write the cells of column A into one text file, the cells of column B into another text file, and
so on"""

# openpyxl modeule helps to deal with excel files
import openpyxl

# Method to convert worksheet into text files
def toTextFiles(filename):
    # load_workbook( ) is used when you have to access an MS Excel file in openpyxl module for some operation
    wb = openpyxl.load_workbook(filename)

    # select inputfile.xlsx(grabs the active worksheet)
    sheet = wb.active
    count = 1

    for colObj in sheet.columns:
        # To create txt files like text-1.txt,etc.
        with open('text-' + str(count) + '.txt', 'w') as file:
            for cellObj in colObj:
                # To write each cell value of respective column in txt file
                file.write(cellObj.value)

        count += 1


# Main method
if __name__ == "__main__":
    # Provided input excel file to convert into text files
    toTextFiles('inputfile.xlsx')
