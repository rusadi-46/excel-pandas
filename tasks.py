import pandas
import numpy as np

# Read file excel and collect data with sheetnama rptOmzet, get col from A-H except D and skip 2 frist row
object_data = pandas.read_excel('omzet.xlsx', sheet_name='rptOmzet', usecols='A:B, E:H', skiprows=2)
# Get index data with non NaN value
index_names = object_data.Employee.dropna().index
# Initiate last row data
last_row = object_data.Employee.index[-1]

# Loop index names to get data omzet per employee
for index in range(len(index_names)):
    last_idx = index_names[-1]
    target_idx = index_names[index]

    if target_idx == last_idx:
        # get data omzet each employee in last index of employee from range row and range column
        omzet_employees = object_data.iloc[target_idx:last_row, 0:7]
    else:
        # get data omzet each employe form range row and range column
        next_index = index_names[index + 1]
        omzet_employees = object_data.iloc[target_idx:next_index, 0:7]
    
    # Initiate employee name conver to title case 
    sheet_name = (omzet_employees['Employee'].values[0]).title()
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
        service_name = CleanServiceName(omzet.Description[target])

        # Rename NaN value with service name
        if target == last:
            omzet.loc[target:last_row_product, 'Description'] = service_name
        else:
            next = index_product[idx + 1] - 1
            omzet.loc[target:next, 'Description'] = service_name

    #  Remove parent service name
    omzet.drop(index_product, axis=0, inplace=True)
    # Sort data ascending by column description
    omzet.sort_values(by=['Description'], ascending=True, inplace=True)

    # write to excel with each sheetname by employee name
    WriteToExcel(omzet, sheet_name)
