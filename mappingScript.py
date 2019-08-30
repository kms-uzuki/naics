# Script created in order to streamline quality control checking of CEUS data.
# Only works with specific files, but you can probably understand what's going on here.
# Far less optimized than it could be
# TODO: delete global dependencies lol

import openpyxl
import time
from collections import defaultdict

start1 = time.time()

book_1 = openpyxl.load_workbook(filename='C:\\Users\\Ksuzuki\\Desktop\\naics_13k 08-19-2019 ALL 4000 - Copy.xlsx')
sheet_1 = book_1.active

book_2 = openpyxl.load_workbook(filename='C:\\Users\\Ksuzuki\\Desktop\\NAICS to SECTOR 2018.xlsx')
sheet_2 = book_2['Electricity 2017']


def cat_assign(beg, end):
    global sheet_1, sheet_2
    map_dict = defaultdict(str)


    for i in range(93, 1038):
        if str(sheet_2.cell(row = i, column = 6).value) == 'Commercial':
            b = sheet_2.cell(row = i, column = 2).value
            h = str(sheet_2.cell(row = i, column = 8).value)
            map_dict[b] = h


    keyy = {
        'OFFICE': "Office",
        'RESTAURANT': "Restaurant",
        'FOOD/LIQUOR': "Food Stores",
        'RETAIL STORE': "Retail",
        'WAREHOUSE': "Warehouse",
        'HEALTH CARE': "Health Care",
        'HOTEL': "Lodging",
        'REFR WAREHOUSE': "Refrigerated Warehouse",
        'COLLEGE': "College",
        'SCHOOL': "School",
        'MISC': "Miscellaneous"
        }
    for i in range(beg, end + 1):
        if sheet_1.cell(i, 31).value == None:
            temp = map_dict[str(sheet_1.cell(i, 27).value)]
            sheet_1.cell(i, 33).value = keyy.get(temp)
        else:
            if str(sheet_1.cell(i, 31).value).upper() == 'INC':
                t = map_dict[str(sheet_1.cell(i, 24).value)]
                sheet_1.cell(i, 33).value = keyy.get(t)
            else:
                temp = map_dict[str(sheet_1.cell(i, 32).value)]
                sheet_1.cell(i, 33).value = keyy.get(temp)
    
    book_1.save(filename='C:\\Users\\Ksuzuki\\Desktop\\naics_13k 08-19-2019 ALL 4000 - Copy.xlsx')

def colorer(begin, end): # this function is spaghetti
    global sheet_1, sheet_2

    for i in range(begin, end + 1):

        _cell = "AG{}".format(i)
        # non-commercial codes
        if str(sheet_1.cell(i, 31).value).lower() != "inc" and sheet_1.cell(i, 33).value == None:
            sheet_1[_cell].fill = openpyxl.styles.GradientFill(stop = ("00FFE4", "FF00E5"))

        # inc, no utility code
        elif str(sheet_1.cell(i, 31).value).lower() == "inc" and sheet_1.cell(i, 33).value == None:
            sheet_1[_cell].fill = openpyxl.styles.PatternFill(start_color = "FF0000", end_color = "FF0000", fill_type = "solid")

        # 1 codes, not done; includes inc with utility codes
        elif sheet_1.cell(i, 31).value == None or str(sheet_1.cell(i, 31).value).lower() == "inc":
            if sheet_1.cell(i, 33).value in sheet_1.cell(i, 10).value: # our building type matches ADM's
                if "Office" in sheet_1.cell(i, 33).value: # office buildings get special color for match
                    sheet_1[_cell].fill = openpyxl.styles.PatternFill(start_color = "F79646", end_color = "F79646", fill_type = "solid")
                else: # match between surveyor's code and ADM's code
                    sheet_1[_cell].fill = openpyxl.styles.PatternFill(start_color = "76933C", end_color = "76933C", fill_type = "solid")
            else: #our building type doesn't match ADM's
                sheet_1[_cell].fill = openpyxl.styles.PatternFill(start_color = "C4BD97", end_color = "C4BD97", fill_type = "solid")

        # all done codes
        elif sheet_1.cell(i, 31).value != None:
            if "Office" == str(sheet_1.cell(i, 33).value) and str(sheet_1.cell(i, 33).value) in str(sheet_1.cell(i, 10).value): # offices match
                sheet_1[_cell].fill = openpyxl.styles.PatternFill(start_color = "00FAFF", end_color = "00FAFF", fill_type = "solid")
            else:
                if sheet_1.cell(i, 33).value != None and str(sheet_1.cell(i, 33).value) not in str(sheet_1.cell(i, 10).value): # building type mismatch
                    sheet_1[_cell].fill = openpyxl.styles.PatternFill(start_color = "FFE500", end_color = "FFE500", fill_type = "solid")
                elif str(sheet_1.cell(i, 33).value) == str(sheet_1.cell(i, 10).value): # match between our code and ADM's code
                    sheet_1[_cell].fill = openpyxl.styles.PatternFill(start_color = "00FF00", end_color = "00FF00", fill_type = "solid")
    book_1.save(filename='C:\\Users\\Ksuzuki\\Desktop\\naics_13k 08-19-2019 ALL 4000 - Copy.xlsx')


if __name__ == "__main__":
    start2 = time.time()
    cat_assign(2, 13184)
    start3 = time.time()
    colorer(2, 13184)

    end = time.time()
    
    elapsed1 = end - start1
    elapsed2 = start3 - start2
    elapsed3 = end - start3
    print("Finished\nTotal elapsed: {}\nFn 1 elapsed: {}\nFn 2 elapsed: {}".format(elapsed1, elapsed2, elapsed3))
