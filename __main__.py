''' 
Making proprietary xlsx formats more accessible.

D. Ellis 2022
'''
# pip install openpyxl pandas
from openpyxl import load_workbook
import argparse, pathlib
import pandas as pd

parser = argparse.ArgumentParser(prog = 'excel2csv', description = 'Convert Excel files to CSV files.')
parser.add_argument( 'fin',nargs=1, type=pathlib.Path, help='Input xlsx file')#argparse.FileType('r')
parser.add_argument('-o','--out', nargs='?', type=str, default='./dataOUT',help='Output Directory') #sys.stdout)


args = parser.parse_args()

try:os.mkdir(args.out)
except:pass

src_file = args.fin[0].__str__()

# Load the spreadsheet
wb = load_workbook(filename = src_file)

# Get all the sheets
sheets = wb.sheetnames

for sheet in sheets:
    for table in wb[sheet].tables.keys():
        dummy = wb[sheet].tables[table]

        df = pd.DataFrame(wb[sheet][dummy.ref]).apply(lambda x: [y.internal_value for y in x])

        df = df.set_index(0,inplace=False).rename(columns=df.iloc[0], inplace = False).iloc[1:]

        df.to_csv(args.out + '/' + sheet + '_' + table + '.csv')

        print(sheet, table)


print('Finished: ' + args.out)




