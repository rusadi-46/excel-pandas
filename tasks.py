import pandas
import numpy as np
from helper import WriteToExcel, CleanServiceName, Employee, Commission

# Read file excel and collect data with sheetnama rptOmzet, get col from A-H except D and skip 2 frist row
object_data = pandas.read_excel('omzet.xlsx', sheet_name='rptOmzet', usecols='A:B, E:H', skiprows=2)
# Get index data with non NaN value
index_names = object_data.Employee.dropna().index
# Initiate last row data
last_row = object_data.Employee.index[-1]

# Loop index names to get data omzet per employee
def Calculation(position, omzet, reduction, commission, idx_item, total, price, disc_value):
    if position == 'Stylist':
        if int(disc_value) > 10:
            calculateCommission(omzet, reduction, commission, idx_item, price) 
        else:
            calculateCommission(omzet, reduction, commission, idx_item, total) 
    else:
        if int(disc_value) > 10:
            calCommission(omzet, commission, idx_item, price) 
        else:
            calCommission(omzet, commission, idx_item, total)

def calCommission(omzet, commission, idx_item, nominal):
    omzet.loc[idx_item, 'Nett'] = (nominal * commission) / 100
    omzet.loc[idx_item, '%'] = (omzet.loc[idx_item, 'Nett'] / nominal) * 100

def calculateCommission(omzet, reduction, commission, idx_item, nominal):
    potongan = omzet.loc[idx_item, 'Potongan'] = (nominal * reduction) / 100
    bruto = omzet.loc[idx_item, 'Bruto'] = nominal - potongan
    nett = omzet.loc[idx_item, 'Nett'] = (bruto * commission) / 100
    omzet.loc[idx_item, 'potongan %'] = (potongan / nominal) * 100 
    omzet.loc[idx_item, 'komisi %'] = (nett / bruto) * 100

def generateCommission(position, omzet, index_product, last_row_product, idx, last, target, service_name, category, reduction, commission):
    if target == last:
        omzet.loc[target:last_row_product, 'Description'] = service_name
        omzet.loc[target:last_row_product, 'Category'] = category
        array_omzet = omzet.loc[target:last_row_product].index
        CommissionPerItem(position, omzet, reduction, commission, array_omzet)
    else:
        next = index_product[idx + 1] - 1
        omzet.loc[target:next, 'Description'] = service_name
        omzet.loc[target:next, 'Category'] = category
        array_omzet = omzet.loc[target:next].index
        CommissionPerItem(position, omzet, reduction, commission, array_omzet)

def CommissionPerItem(position, omzet, reduction, commission, array_omzet):
    for idx_item in array_omzet:
        total = omzet.loc[idx_item, 'Total Nett']
        price = omzet.loc[idx_item,'Price']
        discount = omzet.loc[idx_item,'Total Disc.']
        disc_value = (discount / price) * 100
                
        print(f'Total = {total}')
        print(f'Price = {price}')
        print(f'Discount = {discount}')
        print(f'Discount % = {disc_value}')
                
        Calculation(position, omzet, reduction, commission, idx_item, total, price, disc_value)

def getDataOmzet(object_data, index_names, last_row, index, last_idx, target_idx):
    if target_idx == last_idx:
        # get data omzet each employee in last index of employee from range row and range column
        omzet_employees = object_data.iloc[target_idx:last_row, 0:7]
    else:
        # get data omzet each employe form range row and range column
        next_index = index_names[index + 1]
        omzet_employees = object_data.iloc[target_idx:next_index, 0:7]
    return omzet_employees


def regenerateDataOmzet(object_data, index_names, last_row):
    for index in range(len(index_names)):
        last_idx = index_names[-1]
        target_idx = index_names[index]

        omzet_employees = getDataOmzet(object_data, index_names, last_row, index, last_idx, target_idx)
    
        # Initiate employee name convert to title case 
        sheet_name = (omzet_employees['Employee'].values[0]).title()
        position = Employee()[sheet_name.lower()].title()
    
        print(sheet_name)
        print(position)
        print('-------------')
    
        omzet = omzet_employees.iloc[0:omzet_employees.shape[0], 1:omzet_employees.shape[1]]
        # get index service name except NaN value from column Descriptions
        index_product = omzet.Description.dropna().index
        # initiate last index from loop service name
        last_row_product = omzet.Description.index[-1]
        
        # Loop index service name
        for idx in range(len(index_product)):
            last = index_product[-1]
            target = index_product[idx]
            # Clean name of services
            service_name = CleanServiceName(omzet.Description[target])['service_name']
            category = CleanServiceName(omzet.Description[target])['category']
            reduction = Commission()[position.lower()][category]['reduction']
            commission = Commission()[position.lower()][category]['commission']
            
            if sheet_name == 'Saini' and category == 'mp':
                commission = int(50)
        
            # Rename NaN value with service name
            generateCommission(position, omzet, index_product, last_row_product, idx, last, target, service_name, category, reduction, commission)
            print('---------------------')
            
        #  Remove parent service name
        omzet.drop(index_product, axis=0, inplace=True)
        # Sort data ascending by column description
        omzet.sort_values(by=['Description'], ascending=True, inplace=True)

        # write to excel with each sheetname by employee name
        destination =  "/Users/detik/Project/extra/excel-pandas/omzet.xlsx"
        WriteToExcel(destination, omzet, sheet_name, False)
        
# regenerateDataOmzet(object_data, index_names, last_row)

df_omzet = pandas.read_excel('omzet.xlsx', sheet_name=None)
actual_sheets = [sheet for sheet in df_omzet if sheet != 'rptOmzet']

for item in actual_sheets:
    data_sheet = pandas.read_excel('omzet.xlsx', sheet_name=item)
    grouped = data_sheet.groupby('Description').sum()
    grouped.loc['Grand Total'] = grouped.sum()
    # grouped.loc[-1] = item  # adding a row
    # grouped.index = grouped.index + 1
    destination = '/Users/detik/Project/extra/excel-pandas/Final Omzet.xlsx'
    
    WriteToExcel(destination, grouped, item, True)
    
    print(item)
    print('-----------------------')

