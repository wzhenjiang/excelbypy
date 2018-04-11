#! /usr/local/bin/python3

# section for module import
import sys
import re	# to handle regular expression
import openpyxl

# section for class definition

START_ROW = 2
START_COL = 2

SHEET_NAME = 'case.plan'

def transpose(fn):
    titles = []
    items = []
    
    wb = openpyxl.load_workbook(fn)
    sh = wb[SHEET_NAME]

    # find out all titles 
    end_col = START_COL+1
    while True:
        val = sh.cell(START_ROW, end_col).value
        if val:
            titles.extend(val.split('.')[:-1])
            end_col += 1
        else: 
            break

    # find out all items
    end_row = START_ROW + 1
    while True:
        val = sh.cell(end_row, START_COL).value
        if val:
            items.append(val)
            end_row += 1
        else:
            break
    end_row -= 1

    print(f'titles: {titles}')
    print(f'items: {items}')

    newsheet = wb.create_sheet('transposed')
    rolling_row , rolling_col = 2,2
    newsheet.cell(rolling_row,rolling_col).value = 'items'
    newsheet.cell(rolling_row,rolling_col+1).value = 'title'
    newsheet.cell(rolling_row,rolling_col+2).value = 'value'

    # write into the new tab with transposed data, here, title first
    for inx_title in len(titles):
        for inx_item in len(items):
            val = sh.cell(inx_item + START_ROW+1,inx_title + START_COL+1).value
            if str(val).upper() != 'NA':
                rolling_row +=1
                rolling_col = 2
                newsheet.cell(rolling_row,rolling_col).value = titles[inx_title]
                newsheet.cell(rolling_row,rolling_col+1).value = items[inx_item]
                newsheet.cell(rolling_row,rolling_col+2).value = val

    wb.save(fn)
    print(f'{len(titles)*len(items)} cells transposed!')

'''
wb = Workbook()
sheet = wb.active
'''


# here for unit test function
def Unit_Test(fn):
    transpose(fn)

# here for main code, usually call for Unit_Test
if __name__=='__main__':
    if len(sys.argv) < 2:
        print('Usage: transpose fn')
        exit()
    fn = sys.argv[1]
    print(f'Input file: {fn}')
    Unit_Test(fn)


