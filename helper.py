import pandas

def WriteToExcel(data, sheetname):
    with pandas.ExcelWriter(
        "/Users/rusadi/Project/excel-python/omzet.xlsx",
        mode="a",
        engine="openpyxl",
        if_sheet_exists="replace",
    ) as writer:
        data.to_excel(writer, sheet_name=sheetname, index=False) 

def CleanServiceName(argument):
    service_name = argument.title()

    if 'blow' in argument.lower():
        service_name = 'Blow'
    if 'cut woman' in argument.lower():
        service_name = 'Haircut Woman'
    if 'cut man' in argument.lower():
        service_name = 'Haircut Man'
    if 'poni' in argument.lower():
        service_name = 'Haircut Bangs'
    if 'inoa' in argument.lower():
        service_name = 'Coloring Inoa'
    if 'fashion color' in argument.lower():
        service_name = 'Fashion Color'
    if 'color' in argument.lower():
        service_name = 'Coloring'
    if 'oleoshape' in argument.lower():
        service_name = 'Oleoshape'
    if 'keratin' in argument.lower():
        service_name = 'Coco Keratin'
    if 'dry' in argument.lower():
        service_name = 'Dry Only'
    if 'repair full' in argument.lower():
        service_name = 'Dry Only'
    if 'sanggul' in argument.lower():
        service_name = 'Upstyle'
    if 'upstyle' in argument.lower():
        service_name = 'Upstyle'
    if 'smoothing' in argument.lower():
        service_name = 'Smoothing'
    if 'higlight' in argument.lower():
        service_name = 'Highlight'
    if 'highlight' in argument.lower():
        service_name = 'Highlight'
    if 'extra loreal 1' in argument.lower():
        service_name = 'Extra Loreal 1'
    if 'extra loreal 1/2' in argument.lower():
        service_name = 'Extra Loreal 1/2'
    if 'shampo' in argument.lower():
        service_name = 'Shampoo'
    if 'shampo naturica' in argument.lower():
        service_name = 'Shampoo Naturica'
    if 'netral' in argument.lower():
        service_name = 'Netral'
    if 'back massage' in argument.lower():
        service_name = 'Back Massage'
    if 'hairlos' in argument.lower():
        service_name = 'Treatment Hairloss'
    if 'cr hair & scalp complete' in argument.lower():
        service_name = 'Treatment Complete'
    if 'dd cream pemakian' in argument.lower():
        service_name = 'Cream DD Pemakaian'
    if 'hair tonic' in argument.lower():
        service_name = 'Hair Tonic'
    if 'hairmask repair expert ' in argument.lower():
        service_name = 'Hairmask Repair Expert'
    if 'hair spa' in argument.lower():
        service_name = 'Loreal Hairspa'
    if 'hairspa' in argument.lower():
        service_name = 'Loreal Hairspa'
    if 'hair spa' in argument.lower():
        service_name = 'Loreal Hairspa'
    if 'naturica' in argument.lower():
        service_name = 'Treatment Naturica'
    if 'cromearth' in argument.lower():
        service_name = 'Naturica Mask Cromearth'
    if 'repairing deep shm 250ml ' in argument.lower():
        service_name = 'Repairing Deep Shampoo 250 ml'
    if 'naturica repairing shm 250ml' in argument.lower():
        service_name = 'Naturica Repairing Shampoo 250 ml'
    if 'repairing shm 250ml' in argument.lower():
        service_name = 'Naturica Repairing Shampoo 250 ml'
    if 'repairing deep shampo 1000ml' in argument.lower():
        service_name = 'Repairing Deep Shampo 1000 ml'
    if 'foot polish' in argument.lower():
        service_name = 'Foot Polish OPI'
    if 'half leg rica wax' in argument.lower():
        service_name = 'Half Leg Rica Wax'
    if 'hand polish' in argument.lower():
        service_name = 'Hand Polish OPI'
    if 'manicure' in argument.lower():
        service_name = 'Manicure Gehwol'
    if 'pedicure' in argument.lower():
        service_name = 'Pedicure Gehwol'
    if 'scrub foot' in argument.lower():
        service_name = 'Foot Scrub'
    if 'scrub foot' in argument.lower():
        service_name = 'Hand Scrub'
    if 'dandruf' in argument.lower():
        service_name = 'Treatment Dandruff'
    if 'foot polish' in argument.lower():
        service_name = 'Foot Polish'
    if 'reflexy' in argument.lower():
        service_name = 'Foot Reflexy'
    if 'reflexi' in argument.lower():
        service_name = 'Foot Reflexy'

    return service_name
