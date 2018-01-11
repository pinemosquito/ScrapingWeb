import win32com.client
from win32com.client import DispatchEx
from datetime import datetime, timezone, timedelta
from pytz import timezone

from bs4 import BeautifulSoup
import requests
import pandas as pd
import numpy as np

ofile = r"C:\Users\xyzabc\Desktop\Summary.xlsx"

now = datetime.now(timezone('US/Pacific'))
output_file = ofile.rsplit('.')[0] + '_' + now.strftime("%m-%d-%Y_%H%M%S") + '.' + ofile.rsplit('.')[1]

BASE_URL = [
'https://www.reuters.com/finance/stocks/company-officers/FB.O',
'http://www.reuters.com/finance/stocks/company-officers/GOOG.O',
'http://www.reuters.com/finance/stocks/company-officers/AMZN',
'http://www.reuters.com/finance/stocks/company-officers/AAPL'
]

# Loading empty array for board members
board_members = []
# Loop through URL
for b in BASE_URL:
    html = requests.get(b).text
    soup = BeautifulSoup(html, "html.parser")
    officer_table = soup.find('table', {"class" : "dataTable"})

    try:
        #loop through table, grab each of the 4 columns shown (try one of the links yourself to see the layout)
        for row in officer_table.find_all('tr'):
            cols = row.find_all('td')
            if len(cols) == 4:
                board_members.append((b, cols[0].text.strip(), cols[1].text.strip(), cols[2].text.strip(), cols[3].text.strip()))
    except: pass

# Convert output to new array and check length
board_array = np.asarray(board_members)
len(board_array)

# Convert new array to dataframe
df = pd.DataFrame(board_array)
df.columns = ['URL', 'Name', 'Age','Year_Joined', 'Title']

# Save to Excel
df.to_excel(output_file, index=False)
print('Results saved to {}'.format(output_file))

xel = DispatchEx("Excel.Application")
workbook = xel.Workbooks.Open(output_file)
worksheet = workbook.Worksheets('Sheet1')
workbook.Close(SaveChanges=1)
xel.Quit
del xel
