# Writing to an excel  
# sheet using Python 
import xlwt 
from xlwt import Workbook 
from datetime import datetime
from os.path import expanduser
home = expanduser("~")

# Workbook is created 
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Composites',cell_overwrite_ok=True) 

styleTitle = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')


sheet1.write(0, 0, 'Composite File Main Path',styleTitle) 
sheet1.write(0, 1, 'Composite Name / Rev',styleTitle) 
sheet1.write(0, 2, 'Relational Type',styleTitle) 
sheet1.write(0, 3, 'Integration Type',styleTitle) 
sheet1.write(0, 4, 'Name',styleTitle) 
sheet1.write(0, 5, 'Location',styleTitle) 
sheet1.write(0, 6, 'Address Info',styleTitle) 


def writetoExcel(rowNub,columnNumb,data):
    """
        This method is used to write Worksheets in a Workbook.

        :param rowNub:
           new item nuber (int)
        :param columnNumb:
           0, 'Composite File Main Path'
           1, 'Composite Name / Rev'
           2, 'Relational Type'
           3, 'Integration Type' 
           4, 'Name'
           5, 'Location'
           6, 'Address Info'

        :param data:
          object data for composite
        """
    sheet1.write(rowNub,columnNumb, data) 

def saveExcel():
    logDate = 'Current date/time: {}'.format( datetime.now() )
    now = datetime.now()
    print(logDate,"\nExport File Path: ",home,r'\Desktop\CompositeRef'+now.strftime("%d%m%Y%H%M%S")+'.xls')
    wb.save(home+r'\Desktop\CompositeRef'+now.strftime("%d%m%Y%H%M%S")+'.xls') 


