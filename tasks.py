import pandas
import numpy as np
from datetime import datetime, timedelta
from helper import WriteToExcel, CleanServiceName, Employee, Commission
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

# Read file excel and collect data with sheetnama rptOmzet, get col from A-H except D and skip 2 frist row
object_data = pandas.read_excel('omzet.xlsx', sheet_name='rptOmzet', usecols='A:B, E:H', skiprows=2)
# Get index data with non NaN value
index_names = object_data.Employee.dropna().index
# Initiate last row data
last_row = object_data.Employee.index[-1]

writer = pandas.ExcelWriter('/Users/rusadi/Projects/extra/excel-pandas/Final Omzet.xlsx', engine='xlsxwriter')

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
        CommissioerItem(position, omzet, reduction, commission, array_omzet)
    else:
        next = index_product[idx + 1] - 1
        omzet.loc[target:next, 'Description'] = service_name
        omzet.loc[target:next, 'Category'] = category
        array_omzet = omzet.loc[target:next].index
        CommissioerItem(position, omzet, reduction, commission, array_omzet)

def CommissioerItem(position, omzet, reduction, commission, array_omzet):
    for idx_item in array_omzet:
        total = omzet.loc[idx_item, 'Total Nett']
        price = omzet.loc[idx_item,'Price']
        discount = omzet.loc[idx_item,'Total Disc.']
        disc_value = (discount / price) * 100
                
        Calculation(position, omzet, reduction, commission, idx_item, total, price, disc_value)

def getDataOmzet(object_data, index_names, last_row, index, last_idx, target_idx):
    if target_idx == last_idx:
        # get data omzet each employee in last index of employee from range row and range column
        omzet_employees = object_data.iloc[target_idx:last_row + 1, 0:7]
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
            category = CleanServiceName(omzet.Description[target])['category'].lower()
            reduction = Commission()[position.lower()][category]['reduction']
            commission = Commission()[position.lower()][category]['commission']
            
            if sheet_name == 'Saini' and category == 'mp':
                commission = int(50)
                
            if sheet_name == 'Marni' and category == 'mp':
                commission = int(25)
                
            if sheet_name == 'Febriana' and category == 'mp':
                commission = int(25)
                
            if sheet_name == 'Tatang' and (category == 'blow' or category == 'haircut'):
                reduction = int(20)
            
            if sheet_name == 'Tatang' and category == 'chemical':
                reduction = int(40)
            
            if sheet_name == 'Nia ' and category == 'makeup':
                commission = int(40)
        
            # Rename NaN value with service name
            generateCommission(position, omzet, index_product, last_row_product, idx, last, target, service_name, category, reduction, commission)
            
        #  Remove parent service name
        omzet.drop(index_product, axis=0, inplace=True)
        # Sort data ascending by column description
        omzet.sort_values(by=['Description'], ascending=True, inplace=True)
        # write to excel with each sheetname by employee name
        destination =  "/Users/rusadi/Projects/extra/excel-pandas/omzet.xlsx"
        WriteToExcel(destination, omzet, sheet_name, False)
        
def generateFinalOmzet():
    df_omzet = pandas.read_excel('omzet.xlsx', sheet_name=None)
    actual_sheets = [sheet for sheet in df_omzet if sheet != 'rptOmzet']
    data_omzet = pandas.read_excel('Payroll.xlsx', sheet_name='Omzet', usecols=['Nama','Omzet','Komisi', 'Bonus Omzet'])
    data_omzet.rename(index=data_omzet.Nama, inplace=True)
    total_omzet = 0
    
    for item in actual_sheets:
        data_sheet = pandas.read_excel('omzet.xlsx', sheet_name=item)

        grouped = data_sheet.groupby('Description').sum()
        grouped.loc['Grand Total'] = grouped.sum()
        grouped.insert(0,'Description', grouped.index)
        
        print(grouped.columns.values)

        if 'potongan %' in grouped.columns.values:
            grouped.drop(columns=['potongan %', 'komisi %', 'Category'], inplace=True)
            grouped.rename(columns = {'Description':'SERVICE', 'Price':'PRICE', 'Total Qty':'QTY', 'Total Disc.':'DISC', 'Total Nett':'AFTER DISC', 'Potongan':'POTONGAN', 'Bruto':'BRUTO', 'Nett':'NETT'}, inplace = True)
        else:
            grouped.drop(columns=['%', 'Category'], inplace=True)
            grouped.rename(columns = {'Description':'SERVICE', 'Price':'PRICE', 'Total Qty':'QTY', 'Total Disc.':'DISC', 'Total Nett':'BRUTO', 'Nett':'NETT'}, inplace = True)
            
        total_omzet = grouped.values[-1][4]
        total_komisi = grouped.values[-1][-1]
        data_omzet.loc[item.strip(), 'Omzet'] = total_omzet
        data_omzet.loc[item.strip(), 'Komisi'] = total_komisi
        
        # Create a Pandas Excel writer using XlsxWriter engine.
        grouped.to_excel(writer, sheet_name=item, startrow=4, index=False)
        last_index =  grouped.shape[0] + 4
        
        # Get workbook and worksheet objects
        workbook  = writer.book
        worksheet = writer.sheets[item]
        max_col = grouped.shape[1]
        
        # Format cell
        merge_format = workbook.add_format({'bold': True, 'align':'center'})
        center_format = workbook.add_format({'align':'center'})
        format_number = workbook.add_format({'num_format': '#,##0'})

        # Set merge_range by length of colums names
        len_cols = len(grouped.columns)
        current_month = datetime.now().strftime("%B")
        previous_month = (datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0) - timedelta(days=1)).strftime("%B")
        year = datetime.now().year
        
        # Merge row title
        worksheet.merge_range(0, 0, 0, len_cols - 1, 'T-Style Salon'.upper(), merge_format)
        worksheet.merge_range(1, 0, 1, len_cols - 1, f'Periode 21 { previous_month } - 20 { current_month } { year }'.upper(), merge_format)
        worksheet.merge_range(2, 0, 2, len_cols - 1, item.upper(), merge_format)
        
        # Format column
        worksheet.set_row(4, None, center_format)
        worksheet.set_column('A:A',30)
        worksheet.set_column(f'B6:B{last_index}',10,format_number)
        worksheet.set_column(f'D6:H{last_index}',10, format_number)
        worksheet.set_column(f'C6:C{last_index}',8, center_format)
        
        # Set style on table
        col_settings = [{ 'header': column } for column in grouped.columns]    
        worksheet.add_table(4, 0, last_index, max_col - 1, { 'autofilter': False, 'columns': col_settings, 'style': 'Table Style Light 11', 'total_row': True })
    
    writer.close()
    
def GenerateAbsensi():
    df_attendance = pandas.read_excel('Absensi.xlsx', sheet_name='absen', skiprows=1, usecols='B:B, AI:AI, AN:AP, AR:AS')
    df_attendance.rename(columns = {'Unnamed: 1':'Name', 'TH':'Total Hadir', 'S':'Sakit', 'I':'Izin', 'C':'Cuti', 'L':'Lembur', 'T':'Telat'}, inplace = True)
    df_attendance['Prestasi Absensi'] = 0
    df_attendance['Potongan'] = 0
    df_attendance['Kasbon'] = 0
    
    css_alt_rows = 'background-color: powderblue; color: black;'
    css_indexes = 'background-color: steelblue; color: white; text-align: center'
    styled = df_attendance.style.apply(lambda col: np.where(col.index % 2, css_alt_rows, None)).applymap_index(lambda _: css_indexes, axis=0).applymap_index(lambda _: css_indexes, axis=1).set_properties(subset=['Total Hadir','Sakit','Izin','Cuti','Lembur','Telat','Prestasi Absensi','Potongan','Kasbon'], **{'text-align': 'center'})
    destination =  "/Users/rusadi/Projects/extra/excel-pandas/Payroll.xlsx"
    WriteToExcel(destination, styled, sheetname="Absensi", index=False)
        
    print("Generate absensi successfuly")
    print("-" * 30)
    
# regenerateDataOmzet(object_data, index_names, last_row)
generateFinalOmzet()
# GenerateAbsensi()