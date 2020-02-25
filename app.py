import xlrd
import sys
from collections import OrderedDict
import simplejson as json
sheet = sys.argv[0]
j = None
# Open the workbook and select the first worksheet
try:
    wb = xlrd.open_workbook(sheet if len(sys.argv) is not 1 else 'example.xlsx')
except:
    print("=== ERROR LOADING XLS/X FILE, CHECK FILE PATH ===")
    exit(0)

try:
    sh = wb.sheet_by_index(0)
    # List to hold dictionaries
    rows = []
    # Iterate through each row in worksheet and fetch values into dict
    titles = sh.row_values(0)

    for rownum in range(1, sh.nrows):
        single_row = OrderedDict()
        row_values = sh.row_values(rownum)
        print row_values
        index = 0
        for title in titles:
            title_key = title.lower().replace(' ', '_')
            single_row[title_key] = row_values[index]
            index = index+1
        rows.append(single_row)
    # Serialize the list of dicts to JSON
    j = json.dumps(rows)
except:
    print("=== Something went wrong while parsing the file ===")
    exit(0)

try:
    # Write to file
    with open('data.json', 'w') as f:
        f.write(j)
except:
    print("=== Something went wrong while writing the file ===")
    exit(0)

