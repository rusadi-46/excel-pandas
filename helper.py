import pandas

def WriteToExcel(destination, data, sheetname, index):
    with pandas.ExcelWriter(
        destination,
        mode="a",
        engine="openpyxl",
        if_sheet_exists="replace",
    ) as writer:
        data.to_excel(writer, sheet_name=sheetname, index=index) 

def CleanServiceName(argument):
    unused_name  = {'(new)', '(N)', '(baru)', '(baru )', '( baru )', 'New', '( baru)'}
    checked      = any(element in argument for element in unused_name)
    service_name = ' '.join(argument.split(' ')[:-1]) if checked else argument.title()
    category     = ''

    if 'back massage' in service_name.lower():
        service_name = 'Back Massage'
        category = 'massage'
    elif 'blow' in service_name.lower():
        service_name = 'Blow'
        category = 'blow'
    elif 'blow extention' in service_name.lower():
        service_name = 'Blow Extention'
        category = 'blow'
    elif 'coco keratin l' in service_name.lower():
        service_name = 'Coco Keratin Long'
        category = 'chemical'
    elif 'color l ' in service_name.lower():
        service_name = 'Coloring Long '
        category = 'chemical'
    elif 'coloring l' in service_name.lower():
        service_name = 'Coloring Long'
        category = 'chemical'
    elif 'coloring s' in service_name.lower():
        service_name = 'Coloring Short'
        category = 'chemical'
    elif 'cr anti hairloss' in service_name.lower():
        service_name = 'Treatment Anti Hairloss'
        category = 'treatment'
    elif 'crimbath complite' in service_name.lower():
        service_name = 'Treatment Complete'
        category = 'treatment'
    elif 'dd cream pemakian' in service_name.lower():
        service_name = 'DD Cream Pemakaian'
        category = 'product'
    elif 'dry only' in service_name.lower():
        service_name = 'Dry Only'
        category = 'blow'
    elif 'energing pemakian' in service_name.lower():
        service_name = 'Energing Pemakaian'
        category = 'product'
    elif 'Fashion Color' in service_name.lower():
        service_name = 'Fashion Color'
        category = 'chemical'
    elif 'hair cut women' in service_name.lower():
        service_name = 'Haircut Women'
        category = 'haircut'
    elif 'hair spa' in service_name.lower():
        service_name = 'Hairspa'
        category = 'treatment'
    elif 'hair tonic' in service_name.lower():
        service_name = 'Hair Tonic'
        category = 'product'
    elif 'moisturazing shampo 250 ml' in service_name.lower():
        service_name = 'Moisturazing Shampoo 250ml'
        category = 'product'
    elif 'naturica energizing treatment 100ml' in service_name.lower():
        service_name = 'Naturica Energizing Treatment 100ml'
        category = 'product'
    elif 'naturica treatment' in service_name.lower():
        service_name = 'Naturica Treatment'
        category = 'treatment'
    elif 'naturica' in service_name.lower():
        service_name = 'Naturica Treatment'
        category = 'treatment'
    elif 'netral' in service_name.lower():
        service_name = 'Netral'
        category = 'shampoo'
    elif 'opuntia oil pemakaian' in service_name.lower():
        service_name = 'Opuntia Oil Pemakaian'
        category = 'product'
    elif 'reflexi' in service_name.lower():
        service_name = 'Foot Reflexy'
        category = 'massage'
    elif 'repair full' in service_name.lower():
        service_name = 'Repair Full'
        category = 'extention'
    elif 'shampoo' in service_name.lower():
        service_name = 'Shampoo'
        category = 'shampoo'
    elif 'treatment anti hairlos' in service_name.lower():
        service_name = 'Treatment Anti Hairloss'
        category = 'treatment'
    elif 'vitamin' in service_name.lower():
        service_name = 'Vitamin'
        category = 'product'
    elif 'anti dendruf' in service_name.lower():
        service_name = 'Treatment Anti Dandruf'
        category = 'treatment'
    elif 'cr hair & scalp' in service_name.lower():
        service_name = 'Treatment Complete'
        category = 'treatment'
    elif 'crimbath complite' in service_name.lower():
        service_name = 'Treatment Complete'
        category = 'treatment'
    elif 'foot polish opi' in service_name.lower():
        service_name = 'Foot Polish OPI'
        category = 'polish'
    elif 'hair spa' in service_name.lower():
        service_name = 'Hairspa'
        category = 'treatment'
    elif 'half leg rica wax' in service_name.lower():
        service_name = 'Half Leg Rica Wax'
        category = 'polish'
    elif 'hand polis opi' in service_name.lower():
        service_name = 'Hand Polis OPI'
        category = 'polish'
    elif 'manicure' in service_name.lower():
        service_name = 'Manicure Gehwol'
        category = 'mp'
    elif 'naturica detox' in service_name.lower():
        service_name = 'Naturica Detox'
        category = 'treatment'
    elif 'naturica energizing tratement 100ml' in service_name.lower():
        service_name = 'Naturica Energizing 100ml'
        category = 'product'
    elif 'netral' in service_name.lower():
        service_name = 'Netral'
        category = 'shampoo'
    elif 'pedicure' in service_name.lower():
        service_name = 'Pedicure Gehwol'
        category = 'mp'
    elif 'reflexy' in service_name.lower():
        service_name = 'Foot Reflexy'
        category = 'massage'
    elif 'soothing shampo 1000 ml' in service_name.lower():
        service_name = 'Shoothing Shampo 1000ml'
        category = 'product'
    elif 'foot gell opie' in service_name.lower():
        service_name = 'Foot Gell OPI'
        category = 'polish'
    elif 'hair tonic' in service_name.lower():
        service_name = 'Hair Tonic'
        category = 'product'
    elif 'hairspa' in service_name.lower():
        service_name = 'Hairspa'
        category = 'treatment'
    elif 'blow' in service_name.lower():
        service_name = 'Blow'
        category = 'blow'
    elif 'coloring inoa s' in service_name.lower():
        service_name = 'Coloring Inoa Short'
        category = 'chemical'
    elif 'coloring l 1' in service_name.lower():
        service_name = 'Coloring Long'
        category = 'chemical'
    elif 'cut man' in service_name.lower():
        service_name = 'Haircut Man'
        category = 'haircut'
    elif 'cut women' in service_name.lower():
        service_name = 'Haircut Woman'
        category = 'haircut'
    elif 'extra loreal 1' in service_name.lower():
        service_name = 'Extra Loreal 1'
        category = 'product'
    elif 'extra loreal 1/2' in service_name.lower():
        service_name = 'Extra Loreal 1/2'
        category = 'product'
    elif 'fashion color' in service_name.lower():
        service_name = 'Fashion Color'
        category = 'chemical'
    elif 'hair coloring long' in service_name.lower():
        service_name = 'Coloring Long'
        category = 'chemical'
    elif 'haircoloring inoa l' in service_name.lower():
        service_name = 'Coloring Inoa Long'
        category = 'chemical'
    elif 'haircut man' in service_name.lower():
        service_name = 'Haircut Man'
        category = 'haircut'
    elif 'haircut woman' in service_name.lower():
        service_name = 'Haircut Woman'
        category = 'haircut'
    elif 'upstyle' in service_name.lower():
        service_name = 'Upstyle/Sanggul'
        category = 'upstyle'
    elif '(naturica) opuntia oil shampo 250ml' in service_name.lower():
        service_name = 'Opuntia Oil Shampoo 250ml'
        category = 'product'
    elif 'hairmask repair expert ' in service_name.lower():
        service_name = 'Hairmask Repair Expert '
        category = 'treatment'
    elif 'shampo' in service_name.lower():
        service_name = 'Shampoo'
        category = 'shampoo'
    elif 'smoothing l' in service_name.lower():
        service_name = 'Smoothing Long'
        category = 'chemical'
    elif 'dry' in service_name.lower():
        service_name = 'Dry Only'
        category = 'blow'
    elif 'foot polish opie' in service_name.lower():
        service_name = 'Foot Polish OPI'
        category = 'polish'
    elif 'hand polish opi' in service_name.lower():
        service_name = 'Hand Polish OPI'
        category = 'polish'
    elif 'waxing 1/4' in service_name.lower():
        service_name = 'Waxing 1/4'
        category = 'polish'
    elif 'apuntia oil shape matt puty' in service_name.lower():
        service_name = 'Opuntia Oil Shape Matt Putty'
        category = 'product'
    elif 'coco keratin s' in service_name.lower():
        service_name = 'Coco Keratin Short'
        category = 'chemical'
    elif 'color inoa s' in service_name.lower():
        service_name = 'Coloring Inoa Short'
        category = 'chemical'
    elif 'color l' in service_name.lower():
        service_name = 'Coloring Long'
        category = 'chemical'
    elif 'coloring s 1' in service_name.lower():
        service_name = 'Coloring Short'
        category = 'chemical'
    elif 'cut man' in service_name.lower():
        service_name = 'Haircut Man'
        category = 'haircut'
    elif 'energizing miracle tratment sppry 100ml' in service_name.lower():
        service_name = 'Energizing Miracle Spray 100ml'
        category = 'product'
    elif 'extra loreal 1' in service_name.lower():
        service_name = 'Extra Loreal 1'
        category = 'product'
    elif 'hair coloring short' in service_name.lower():
        service_name = 'Coloring Short'
        category = 'chemical'
    elif 'hair cut man' in service_name.lower():
        service_name = 'Haircut Man'
        category = 'haircut'
    elif 'hair cut women' in service_name.lower():
        service_name = 'hair cut women by tatang'
        category = 'haircut'
    elif 'hair extention' in service_name.lower():
        service_name = 'Hair Extention'
        category = 'extention'
    elif 'haircut women' in service_name.lower():
        service_name = 'Haircut Women'
        category = 'haircut'
    elif 'higlight' in service_name.lower():
        service_name = 'Highlight'
        category = 'chemical'
    elif 'repair 1/4' in service_name.lower():
        service_name = 'Repair 1/4'
        category = 'extention'
    elif 'shape matt' in service_name.lower():
        service_name = 'Shape Matt Putty for Men'
        category = 'product'
    elif 'smoothing' in service_name.lower():
        service_name = 'smoothing'
        category = 'chemical'
    elif 'smothing coco' in service_name.lower():
        service_name = 'Smoothing Coco'
        category = 'chemical'

    print(service_name)
    return {
        'service_name':  service_name,
        'category': category
    }

def Commission():
    return {
        'stylist':
        {
            'blow': {
                'reduction': 30,
                'commission': 50
            },
            'chemical': {
                'reduction': 50,
                'commission': 50
            },
            'haircut': {
                'reduction': 30,
                'commission': 50
            },
            'product': {
                'reduction': 0,
                'commission': 10
            },
            'extention': {
                'reduction': 0,
                'commission': 15
            },
            'makeup': {
                'reduction': 0,
                'commission': 50
            },
            'upstyle': {
                'reduction': 0,
                'commission': 40
            },
            'shampoo': {
                'reduction': 0,
                'commission': 50
            }
        },
        'therapies': {
            
            'blow': {
                'reduction': 0,
                'commission': 25
            },
            'chemical': {
                'reduction': 0,
                'commission': 15
            },
            'haircut': {
                'reduction': 0,
                'commission': 20
            },
            'product': {
                'reduction': 0,
                'commission': 10
            },
            'extention': {
                'reduction': 0,
                'commission': 15
            },
            'makeup': {
                'reduction': 0,
                'commission': 50
            },
            'upstyle': {
                'reduction': 0,
                'commission': 40
            },
            'treatment': {
                'reduction': 0,
                'commission': 15
            },
            'mp': {
                'reduction': 0,
                'commission': 20
            },
            'massage': {
                'reduction': 0,
                'commission': 40
            },
            'polish': {
                'reduction': 0,
                'commission': 15
            },
            'shampoo': {
                'reduction': 0,
                'commission': 50
            }
        }
    }

def Employee():
    return {
        'adit': 'therapies',
        'apriyanti': 'therapies',
        'febriana': 'therapies',
        'marni': 'therapies',
        'narko': 'stylist',
        'nuy ': 'therapies',
        'saini': 'therapies',
        'tatang': 'stylist',
        'sinta': 'kasir',
        'dilla': 'kasir'
    }
