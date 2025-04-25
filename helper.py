import pandas

def WriteToExcel(destination, data, sheetname, index=False, startrow=0):
    with pandas.ExcelWriter(
        destination,
        mode="a",
        engine="openpyxl",
        if_sheet_exists="replace",
    ) as writer:
        data.to_excel(writer, sheet_name=sheetname, startrow=startrow, index=index) 

def CleanServiceName(argument):
    unused_name  = {'(new)', '(N)', '(baru)', '(baru )', '( baru )', 'New', '( baru)'}
    checked      = any(element in argument for element in unused_name)
    service_name = ' '.join(argument.split(' ')[:-1]) if checked else argument.title()
    category     = ''
    
    print(service_name)

    if 'back massage' in service_name.lower():
        service_name = 'Back Massage'
        category = 'massage'
    elif 'blow' in service_name.lower():
        service_name = 'Blow'
        category = 'blow'
    elif 'aminexil' in service_name.lower():
        service_name = 'Aminexil'
        category = 'product'
    elif 'aminexi...' in service_name.lower():
        service_name = 'Aminexil'
        category = 'product'
    elif 'member' in service_name.lower():
        service_name = 'Kartu Member'
        category = 'product'
    elif 'blow extention' in service_name.lower():
        service_name = 'Blow Extention'
        category = 'blow'
    elif 'coco keratin l' in service_name.lower():
        service_name = 'Coco Keratin Long'
        category = 'chemical'
    elif 'color l' in service_name.lower():
        service_name = 'Coloring Long'
        category = 'chemical'
    elif 'colour fashion' in service_name.lower():
        service_name = 'Fashion Color'
        category = 'chemical'
    elif 'oleoshape l' in service_name.lower():
        service_name = 'Oleoshape Long'
        category = 'chemical'
    elif 'coloring l' in service_name.lower():
        service_name = 'Coloring Long'
        category = 'chemical'
    elif 'color produk sendiri' in service_name.lower():
        service_name = 'Color Produk Sendiri'
        category = 'chemical'
    elif 'coloring s' in service_name.lower():
        service_name = 'Coloring Short'
        category = 'chemical'
    elif 'cr anti hairloss' in service_name.lower():
        service_name = 'Treatment Anti Hairloss'
        category = 'treatment'
    elif 'anti hairloss' in service_name.lower():
        service_name = 'Treatment Anti Hairloss'
        category = 'treatment'
    elif 'hairmask repair expert' in service_name.lower():
        service_name = 'Hairmask Repair Expert'
        category = 'treatment'
    elif 'Absolut Repair Molecular'.lower() in service_name.lower():
        service_name = 'Absolut Repair Molecular'
        category = 'treatment'
    elif 'Instan Repair Boost'.lower() in service_name.lower():
        service_name = 'Instan Repair Boost'
        category = 'treatment'
    elif 'Absolute Repair Molecular'.lower() in service_name.lower():
        service_name = 'Absolute Repair Molecular'
        category = 'treatment'
    elif 'Absolute Repair Molecular'.lower() in service_name.lower():
        service_name = 'Absolute Repair Molecular'
        category = 'treatment'
    elif 'crimbath complite' in service_name.lower():
        service_name = 'Treatment Complete'
        category = 'treatment'
    elif 'dd cream pemakian' in service_name.lower():
        service_name = 'DD Cream Pemakaian'
        category = 'product'
    elif 'Hair Spray'.lower() in service_name.lower():
        service_name = 'Hair Spray'
        category = 'product'
    elif 'Hair Spray Silhoutte 500Ml'.lower() in service_name.lower():
        service_name = 'Hair Spray Silhoutte 500Ml'
        category = 'product'
    elif 'moist defense mask 250ml' in service_name.lower():
        service_name = 'Moist Defense Mask 250Ml'
        category = 'product'
    elif 'dd cream pemakaian' in service_name.lower():
        service_name = 'DD Cream Pemakaian'
        category = 'product'
    elif 'dry only' in service_name.lower():
        service_name = 'Dry Only'
        category = 'blow'
    elif 'energing pemakian' in service_name.lower():
        service_name = 'Energing Pemakaian'
        category = 'product'
    elif 'root spray pemakian' in service_name.lower():
        service_name = 'Root Spray Pemakaian'
        category = 'product'
    elif 'hair mask color expert' in service_name.lower():
        service_name = 'Hairmask Color Expert'
        category = 'treatment'
    elif 'energizing tratement' in service_name.lower():
        service_name = 'Energizing Tratement'
        category = 'product'
    elif 'Fashion Color' in service_name.lower():
        service_name = 'Fashion Color'
        category = 'chemical'
    elif 'color inoa l' in service_name.lower():
        service_name = 'Coloring Inoa Long'
        category = 'chemical'
    elif 'colour inoa l' in service_name.lower():
        service_name = 'Coloring Inoa Long'
        category = 'chemical'
    elif 'coco keratin' in service_name.lower():
        service_name = 'Coco Keratin'
        category = 'chemical'
    elif 'noni' in service_name.lower():
        service_name = 'Coloring Noni'
        category = 'chemical'
    elif 'color prodact sendiri' in service_name.lower():
        service_name = 'Coloring Produk Sendiri'
        category = 'chemical'
    elif 'hair cut women' in service_name.lower():
        service_name = 'Haircut Woman'
        category = 'haircut'
    elif 'cut woman stylist' in service_name.lower():
        service_name = 'Haircut Woman'
        category = 'haircut'
    elif 'cut poni' in service_name.lower():
        service_name = 'Haircut Bangs'
        category = 'haircut'
    elif 'gunting poni' in service_name.lower():
        service_name = 'Haircut Bangs'
        category = 'haircut'
    elif 'Milk Shake Volumizing Shampoo New'.lower() in service_name.lower():
        service_name = 'Milk Shake Volumizing Shampoo 300ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Whipped Cream 200Ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Whipped Cream 200ml'
        category = 'product_milk_shake'
    elif 'shampo free' in service_name.lower():
        service_name = 'Shampoo'
        category = 'shampoo'
    elif 'prodak tambahan sanggul' in service_name.lower():
        service_name = 'Product Sanggul'
        category = 'product'
    elif 'sanggu tambahan produk' in service_name.lower():
        service_name = 'Product Sanggul'
        category = 'product'
    elif 'apuntia oil tratmen pemakaian' in service_name.lower():
        service_name = 'Opuntia Oil Treatment Pemakaian'
        category = 'product'
    elif 'free' in service_name.lower():
        service_name = 'Service Free'
        category = 'product'
    elif 'hair spa' in service_name.lower():
        service_name = 'Hairspa'
        category = 'treatment'
    elif 'repairing mask 2023' in service_name.lower():
        service_name = 'Naturica Repairing Mask Treatment'
        category = 'treatment'
    elif 'cr complete (cromearth mask)' in service_name.lower():
        service_name = 'Cromearth Mask Treatment'
        category = 'treatment'
    elif 'Hair Mask Cromeath'.lower() in service_name.lower():
        service_name = 'Cromearth Mask Treatment'
        category = 'treatment'
    elif 'hair tonic' in service_name.lower():
        service_name = 'Hair Tonic'
        category = 'product'
    elif 'hair mask repair expert' in service_name.lower():
        service_name = 'Hair Mask Repair Expert'
        category = 'treatment'
    elif 'hair mask moisturazing' in service_name.lower():
        service_name = 'Hair Mask Moisturazing'
        category = 'treatment'
    elif 'moisturazing shampo 250 ml' in service_name.lower():
        service_name = 'Moisturazing Shampoo 250ml'
        category = 'product'
    elif '(NATURICA) Repairing Mask 250ml'.lower() in service_name.lower():
        service_name = 'Naturica Repairing Mask 250ml'
        category = 'product'
    elif 'Naturica Repairing Shm 250ml 2021'.lower() in service_name.lower():
        service_name = 'Naturica Repairing Shampoo 250ml'
        category = 'product'
    elif 'Earth Moon Shampo 250ml'.lower() in service_name.lower():
        service_name = 'Naturica Earth Moon Mask 250ml'
        category = 'product'
    elif 'Naturica Moisturizing defense shmp 1000ml'.lower() in service_name.lower():
        service_name = 'Naturica Moisturizing Defense Shampoo 1000ml'
        category = 'product'
    elif 'naturica earth moon mask 250ml' in service_name.lower():
        service_name = 'Naturica Earth Moon Mask 250ml'
        category = 'product'
    elif 'naturica cromearth earth moon shampoo 250ml' in service_name.lower():
        service_name = 'Naturica Cromearth Moon Shampoo 250ml '
        category = 'product'
    elif 'naturica energizing shampo 1000ml' in service_name.lower():
        service_name = 'Naturica Energizing Shampo 1000ml'
        category = 'product'
    elif 'naturica energizing miracle shampoo 250 ml' in service_name.lower():
        service_name = 'Naturica Energizing Miracle Shampoo 250 ml'
        category = 'product'
    elif 'energizing miracle treatment serum d 100ml' in service_name.lower():
        service_name = 'Energizing Miracle Treatment Serum D 100Ml'
        category = 'product'
    elif 'produk absolut repair masq  200 ml' in service_name.lower():
        service_name = 'Produk Absolut Repair Masq 200ml'
        category = 'product'
    elif 'Opun Tia Oli Dd Hair Crem'.lower() in service_name.lower():
        service_name = 'Opuntia Oil DD Cream 120ml'
        category = 'product'
    elif 'opuntia oil dd cream 120 ml' in service_name.lower():
        service_name = 'Opuntia Oil DD Cream 120ml'
        category = 'product'
    elif 'moist defense cond 200ml' in service_name.lower():
        service_name = 'Moist Defense Conditioner 200ml'
        category = 'product'
    elif 'opuntia oil low shp 1000ml' in service_name.lower():
        service_name = 'Opuntia Oil Low Shp 1000ml'
        category = 'product'
    elif 'opuntia oil shampo 250ml' in service_name.lower():
        service_name = 'Opuntia Oil Shampo 250ml'
        category = 'product'
    elif 'moist defense cond 1000ml' in service_name.lower():
        service_name = 'Moist Defense Cond 1000Ml'
        category = 'product'
    elif 'power mix colour' in service_name.lower():
        service_name = 'Power Mix Colour'
        category = 'product'
    elif 'netral' in service_name.lower():
        service_name = 'Netral'
        category = 'shampoo'
    elif 'naturica opuntia oil conde 200ml' in service_name.lower():
        service_name = 'Naturica Opuntia Oil Conditioner 200ml'
        category = 'product'
    elif 'naturica repairing deep mask 1000 ml' in service_name.lower():
        service_name = 'Naturica Repairing Deep Mask 1000 ml'
        category = 'product'
    elif 'Repairing Deep Shm 1000Ml'.lower() in service_name.lower():
        service_name = 'Naturica Repairing Deep Mask 1000 ml'
        category = 'product'
    elif 'naturica repairing deep shm 250ml' in service_name.lower():
        service_name = 'Naturica Repairing Deep Shm 250ml'
        category = 'product'
    elif 'repairing deep shampo 250' in service_name.lower():
        service_name = 'Naturica Repairing Deep Shm 250ml'
        category = 'product'
    elif 'opuntia oil pemakaian' in service_name.lower():
        service_name = 'Opuntia Oil Pemakaian'
        category = 'product'
    elif 'naturica bamboo detangled hair brush' in service_name.lower():
        service_name = 'Naturica Bamboo Detangled Hair Brush'
        category = 'product'
    elif 'naturica energizing treatment 100ml' in service_name.lower():
        service_name = 'Naturica Energizing Treatment 100ml'
        category = 'product'
    elif 'conditioner naturica' in service_name.lower():
        service_name = 'Conditioner Naturica'
        category = 'product'
    elif 'volumezing shampo 250 ml' in service_name.lower():
        service_name = 'Volumezing Shampo 250 ml'
        category = 'product'
    elif 'soothing shampo 300ml' in service_name.lower():
        service_name = 'Soothing Shampo 300ml'
        category = 'product'
    elif 'METAL DX SHAMPO 300ML'.lower() in service_name.lower():
        service_name = 'Metal DX Shampoo 300ML'
        category = 'product'
    elif 'Milk Shake Leave In Conditioner 350Ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Leave In Conditioner 350Ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Volumizing Shampoo New 300ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Volumizing Shampoo 300ml'
        category = 'product_milk_shake'
    elif 'Volumezing Condetioner 200Ml'.lower() in service_name.lower():
        service_name = 'Volumezing Condetioner 200Ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Daily Frequent Conditioner 300Ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Daily Frequent Conditioner 300Ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Silver Shine Conditioner 250Ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Silver Shine Conditioner 250Ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Moisture Plus Whipped Cream 200Ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Moisture Plus Whipped Cream 200Ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Energizing Conditioner 300Ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Energizing Conditioner 300Ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Conditioning Whipped Cream 200Ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Conditioning Whipped Cream 200Ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Leave In Conditioner 350Ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Leave In Conditioner 350Ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Integrity Intensive Treatment New'.lower() in service_name.lower():
        service_name = 'Milk Shake Integrity Intensive Treatment New'
        category = 'product_milk_shake'
    elif 'Milk Shake Silver Shine Shampoo 300ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Silver Shine Shampoo 300ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Silver Shine Light Shampoo 300ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Silver Shine Light Shampoo 300ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Colour Maintainer Shampoo 300ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Colour Maintainer Shampoo 300ml'
        category = 'product_milk_shake'
    elif 'milk shake incredible milk 150ml' in service_name.lower():
        service_name = 'Milk Shake Incredible Milk 150Ml'
        category = 'product_milk_shake'
    elif 'milk shake moisture plus shampoo 300ml' in service_name.lower():
        service_name = 'Milk Shake Moisture Plus Shampoo 300ml'
        category = 'product_milk_shake'
    elif 'milk shake energizing treatment 30ml' in service_name.lower():
        service_name = 'Milk Shake Energizing Treatment 30Ml'
        category = 'product_milk_shake'
    elif 'Milk Shake Energizing Shampo 300ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Energizing Shampo 300ml'
        category = 'product_milk_shake'
    elif 'naturica treatment' in service_name.lower():
        service_name = 'Naturica Treatment'
        category = 'treatment'
    elif 'Shampo Naturica Pemakian (2023)'.lower() in service_name.lower():
        service_name = 'Shampoo Naturica'
        category = 'product'
    elif 'shampo naturica' in service_name.lower():
        service_name = 'Shampoo Naturica'
        category = 'product'
    elif 'shampoo naturica' in service_name.lower():
        service_name = 'Shampoo Naturica'
        category = 'product'
    elif 'naturica moisturizing conditioner 200ml' in service_name.lower():
        service_name = 'Naturica Moisturizing Conditioner 200ml'
        category = 'product'
    elif 'repairing deep shampo 1000ml' in service_name.lower():
        service_name = 'Naturica Repairing Deep Shampoo 1000ml'
        category = 'product'
    elif 'Terra Cotta Mask 250ml'.lower() in service_name.lower():
        service_name = 'Naturica Terra Cotta Mask 250ml'
        category = 'product'
    elif 'naturica energizing tratement 100ml' in service_name.lower():
        service_name = 'Naturica Energizing 100ml'
        category = 'product'
    elif 'naturica volumizing condetioner 200ml 3' in service_name.lower():
        service_name = 'Naturica Volumizing Conditioner 200ml'
        category = 'product'
    elif 'naturica volumizing condetioner 200ml' in service_name.lower():
        service_name = 'Naturica Volumizing Conditioner 200ml'
        category = 'product'
    elif 'naturica volumizing condetioner 250 ml 2' in service_name.lower():
        service_name = 'Naturica Volumizing Conditioner 250ml'
        category = 'product'
    elif 'Volumezing Condetioner 1000Ml'.lower() in service_name.lower():
        service_name = 'Naturica Volumizing Conditioner 1000Ml'
        category = 'product'
    elif 'naturica volumizing condetioner 250 ml' in service_name.lower():
        service_name = 'Naturica Volumizing Conditioner 250ml'
        category = 'product'
    elif 'naturica volumizing exp condi 200ml 2021' in service_name.lower():
        service_name = 'Naturica Volumizing Exp Condi 200ml'
        category = 'product'
    elif 'naturica volumizing exp condi 200ml' in service_name.lower():
        service_name = 'Naturica Volumizing Exp Condi 200ml'
        category = 'product'
    elif '(NATURICA) Energizing Shampoo 250ml'.lower() in service_name.lower():
        service_name = 'Energizing Shampoo 250ml'
        category = 'product'
    elif 'Naturica Mask Cromearth 2021'.lower() in service_name.lower():
        service_name = 'Naturica Mask Cromearth'
        category = 'product'
    elif '(NATURICA) Balancing Shampoo 1000ml'.lower() in service_name.lower():
        service_name = 'Balancing Shampoo 1000ml'
        category = 'product'
    elif 'Opintial Oil Treatment 50Ml'.lower() in service_name.lower():
        service_name = 'Naturica Opuntial Oil Treatment 50ml'
        category = 'product'
    elif '(NATURICA) Opuntia Oil Treatment 120ml'.lower() in service_name.lower():
        service_name = 'Naturica Opuntial Oil Treatment 120ml'
        category = 'product'
    elif 'Naturica Opuntial Oil Treatment 50ml'.lower() in service_name.lower():
        service_name = 'Naturica Opuntial Oil Treatment 50ml'
        category = 'product'
    elif '(NATURICA) Opuntia Oil DD Haircream 150ml'.lower() in service_name.lower():
        service_name = 'Opuntia Oil DD Haircream 150ml'
        category = 'product'
    elif 'naturica opuntia oil' in service_name.lower():
        service_name = 'Naturica Opuntia Oil'
        category = 'product'
    elif 'naturica shampo volumizing' in service_name.lower():
        service_name = 'Naturica Shampoo Volumizing'
        category = 'product'
    elif 'naturica repairing mask' in service_name.lower():
        service_name = 'Naturica Repairing Mask'
        category = 'treatment'
    elif 'moisturaizing mask 2023' in service_name.lower():
        service_name = 'Naturica Moisturaizing Mask'
        category = 'treatment'
    elif 'naturica repairing deep msk 250ml' in service_name.lower():
        service_name = 'Naturica Repairing Deep Mask 250ml'
        category = 'product'
    elif 'repairing deep mask 250ml' in service_name.lower():
        service_name = 'Naturica Repairing Deep Mask 250ml'
        category = 'product'
    elif 'detoxifing comfort scrub 200ml' in service_name.lower():
        service_name = 'Detoxifing Comfort Scrub 200ml'
        category = 'product'
    elif 'naturica moisturizing condi 200ml 2021' in service_name.lower():
        service_name = 'Naturica Moisturizing Condi 200ml'
        category = 'product'
    elif 'naturica detoxifying shampo 250ml' in service_name.lower():
        service_name = 'Naturica Detoxifying Shampoo 250ml'
        category = 'product'
    elif 'Milk Shake Energizing Shampo 300ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Energizing Shampo 300ml'
        category = 'product_milk_shake'
    elif 'naturica detox' in service_name.lower():
        service_name = 'Naturica Detox Scrub'
        category = 'treatment'
    elif 'naturica' in service_name.lower():
        service_name = 'Naturica Treatment'
        category = 'treatment'
    elif 'reflexi' in service_name.lower():
        service_name = 'Foot Reflexy'
        category = 'massage'
    elif 'balancing shampo 250 ml' in service_name.lower():
        service_name = 'Balancing Remedy Shp 250ml'
        category = 'product'
    elif 'repair full' in service_name.lower():
        service_name = 'Repair Full'
        category = 'extention'
    elif 'Repair New Extention'.lower() in service_name.lower():
        service_name = 'Repair Extention'
        category = 'extention'
    elif 'shampoo' in service_name.lower():
        service_name = 'Shampoo'
        category = 'shampoo'
    elif 'Shampo 4 (New 2023)'.lower() in service_name.lower():
        service_name = 'Shampoo'
        category = 'shampoo'
    elif 'Milk Shake Volumizing Shampoo New 300ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Volumizing Shampoo 300ml'
        category = 'product_milk_shake'
    elif 'treatment anti hairlos' in service_name.lower():
        service_name = 'Treatment Anti Hairloss'
        category = 'treatment'
    elif 'vitamin' in service_name.lower():
        service_name = 'Vitamin'
        category = 'product'
    elif 'treatment anti dandruff' in service_name.lower():
        service_name = 'Treatment Anti Dandruff'
        category = 'treatment'
    elif 'anti dandruf scalp' in service_name.lower():
        service_name = 'Treatment Anti Dandruff'
        category = 'treatment'
    elif 'anti dendruf' in service_name.lower():
        service_name = 'Treatment Anti Dandruff'
        category = 'treatment'
    elif 'cr hair & scalp' in service_name.lower():
        service_name = 'Treatment Complete'
        category = 'treatment'
    elif 'crimbath complite' in service_name.lower():
        service_name = 'Treatment Complete'
        category = 'treatment'
    elif 'nail extention' in service_name.lower():
        service_name = 'Nail Extention'
        category = 'nail_extention'
    elif 'nail art' in service_name.lower():
        service_name = 'Nail Art'
        category = 'nail_art'
    elif 'foot polish opi' in service_name.lower():
        service_name = 'Foot Polish OPI'
        category = 'polish'
    elif 'Hand French Polish OPI'.lower() in service_name.lower():
        service_name = 'Hand French Polish OPI'
        category = 'polish'
    elif 'foot polish' in service_name.lower():
        service_name = 'Foot Polish'
        category = 'polish'
    elif 'hair spa' in service_name.lower():
        service_name = 'Hairspa'
        category = 'treatment'
    elif 'half leg rica wax' in service_name.lower():
        service_name = 'Half Leg Rica Wax'
        category = 'polish'
    elif 'Full Arm Rica Wax'.lower() in service_name.lower():
        service_name = 'Full Arm Rica Wax'
        category = 'polish'
    elif '1/4 waxing' in service_name.lower():
        service_name = 'Waxing 1/4'
        category = 'polish'
    elif 'Hand Waxing'.lower() in service_name.lower():
        service_name = 'Hand Waxing'
        category = 'polish'
    elif 'Foot Waxing Full'.lower() in service_name.lower():
        service_name = 'Foot Waxing'
        category = 'polish'
    elif 'waxing under arm' in service_name.lower():
        service_name = 'Waxing Under Arm'
        category = 'polish'
    elif 'hand polis opi' in service_name.lower():
        service_name = 'Hand Polis OPI'
        category = 'polish'
    elif 'gel hand opi' in service_name.lower():
        service_name = 'Gell Hand OPI'
        category = 'polish'
    elif 'foot french gel opi' in service_name.lower():
        service_name = 'Foot French Gel OPI'
        category = 'polish'
    elif 'hand french gel opi' in service_name.lower():
        service_name = 'Hand French Gel OPI'
        category = 'polish'
    elif 'frech gell hand opie' in service_name.lower():
        service_name = 'Hand French Gel OPI'
        category = 'polish'
    elif 'hand gel remove' in service_name.lower():
        service_name = 'Hand Gell Remove'
        category = 'polish'
    elif 'Hand Remove'.lower() in service_name.lower():
        service_name = 'Hand Gell Remove'
        category = 'polish'
    elif 'remove hand gell' in service_name.lower():
        service_name = 'Hand Gell Remove'
        category = 'polish'
    elif 'Foot Remove'.lower() in service_name.lower():
        service_name = 'Foot Remove'
        category = 'polish'
    elif 'hand gell' in service_name.lower():
        service_name = 'Hand Gell OPI'
        category = 'polish'
    elif 'hand gel' in service_name.lower():
        service_name = 'Hand Gell OPI'
        category = 'polish'
    elif 'remove foot' in service_name.lower():
        service_name = 'Remove Foot Gell'
        category = 'polish'
    elif 'remove hand' in service_name.lower():
        service_name = 'Remove Hand Gell'
        category = 'polish'
    elif 'Remove Ii'.lower() in service_name.lower():
        service_name = 'Remove Gell'
        category = 'polish'
    elif 'Remove I'.lower() in service_name.lower():
        service_name = 'Remove Gell'
        category = 'polish'
    elif 'manicure' in service_name.lower():
        service_name = 'Manicure Gehwol'
        category = 'mp'
    elif 'netral' in service_name.lower():
        service_name = 'Netral'
        category = 'shampoo'
    elif 'pedicure' in service_name.lower():
        service_name = 'Pedicure Gehwol'
        category = 'mp'
    elif 'reflexy' in service_name.lower():
        service_name = 'Foot Reflexy'
        category = 'massage'
    elif 'soothing relief shampo' in service_name.lower():
        service_name = 'Soothing Relief Shampo'
        category = 'product'
    elif 'Framesi Wax 515'.lower() in service_name.lower():
        service_name = 'Framesi Wax 515'
        category = 'product'
    elif 'soothing shampo 1000 ml' in service_name.lower():
        service_name = 'Shoothing Shampo 1000ml'
        category = 'product'
    elif 'pemakaian opuntia oil treatment' in service_name.lower():
        service_name = 'Pemakaian Opuntia Oil Treatment'
        category = 'product'
    elif 'opuntia oil shp 250' in service_name.lower():
        service_name = 'Opuntia Oil Shampoo 250ml'
        category = 'product'
    elif 'volumezing shampo 250' in service_name.lower():
        service_name = 'Naturica Volumezing Shampo 250ml'
        category = 'product'
    elif 'Naturica Volumizing Condetioner 250 Ml 2'.lower() in service_name.lower():
        service_name = 'Naturica Volumizing Condetioner 250 Ml'
        category = 'product'
    elif 'opuntia oil treatment 50 ml' in service_name.lower():
        service_name = 'Opuntia Oil Treatment 50 Ml'
        category = 'product'
    elif 'foot gel remove' in service_name.lower():
        service_name = 'Foot Gell Remove'
        category = 'polish'
    elif 'foot gell opie' in service_name.lower():
        service_name = 'Foot Gell OPI'
        category = 'polish'
    elif 'foot gel' in service_name.lower():
        service_name = 'Foot Gell OPI'
        category = 'polish'
    elif 'hair tonic' in service_name.lower():
        service_name = 'Hair Tonic'
        category = 'product'
    elif 'hairspa ice scrub' in service_name.lower():
        service_name = 'Hairspa Ice Scrub'
        category = 'product'
    elif 'hairspa' in service_name.lower():
        service_name = 'Hairspa'
        category = 'treatment'
    elif 'scalp complete' in service_name.lower():
        service_name = 'Treatment Complete'
        category = 'treatment'
    elif 'blow' in service_name.lower():
        service_name = 'Blow'
        category = 'blow'
    elif 'coloring inoa s' in service_name.lower():
        service_name = 'Coloring Inoa Short'
        category = 'chemical'
    elif 'coloring inoa l' in service_name.lower():
        service_name = 'Coloring Inoa Long'
        category = 'chemical'
    elif 'cleansing' in service_name.lower():
        service_name = 'Cleansing'
        category = 'chemical'
    elif 'coloring l 1' in service_name.lower():
        service_name = 'Coloring Long'
        category = 'chemical'
    elif 'color smartbond long' in service_name.lower():
        service_name = 'Color Smartbond Long'
        category = 'chemical'
    elif 'color smartbond short' in service_name.lower():
        service_name = 'Color Smartbond Short'
        category = 'chemical'
    elif 'coloring' in service_name.lower():
        service_name = 'Coloring'
        category = 'chemical'
    elif 'cut man' in service_name.lower():
        service_name = 'Haircut Man'
        category = 'haircut'
    elif 'cut women' in service_name.lower():
        service_name = 'Haircut Woman'
        category = 'haircut'
    elif 'extra loreal 1/2' in service_name.lower():
        service_name = 'Extra Loreal 1/2'
        category = 'product'
    elif 'extra inoa 1' in service_name.lower():
        service_name = 'Extra Inoa 1'
        category = 'product'
    elif 'balancing remedy shp 1000ml' in service_name.lower():
        service_name = 'Balancing Remedy Shp 1000ml'
        category = 'product'
    elif 'extra inoa 1/2' in service_name.lower():
        service_name = 'Extra Inoa 1/2'
        category = 'product'
    elif 'extra loreal 1' in service_name.lower():
        service_name = 'Extra Loreal 1'
        category = 'product'
    elif 'extra loreal' in service_name.lower():
        service_name = 'Extra Loreal 1'
        category = 'product'
    elif 'fashion color' in service_name.lower():
        service_name = 'Fashion Color'
        category = 'chemical'
    elif 'fashion' in service_name.lower():
        service_name = 'Fashion Color'
        category = 'chemical'
    elif 'highlight' in service_name.lower():
        service_name = 'Highlight'
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
    elif 'Sanggul / Up Style'.lower() in service_name.lower():
        service_name = 'Upstyle/Sanggul'
        category = 'upstyle'
    elif 'styling..' in service_name.lower():
        service_name = 'Upstyle/Sanggul'
        category = 'upstyle'
    elif 'hair style' in service_name.lower():
        service_name = 'Upstyle/Sanggul'
        category = 'upstyle'
    elif '(naturica) opuntia oil shampo 250ml' in service_name.lower():
        service_name = 'Opuntia Oil Shampoo 250ml'
        category = 'product'
    elif 'naturica cromearth earth moon shampoo 250ml' in service_name.lower():
        service_name = 'Naturica Cromearth Earth Moon Shampoo 250ml'
        category = 'product'
    elif 'hairmask repair expert ' in service_name.lower():
        service_name = 'Hairmask Repair Expert'
        category = 'treatment'
    elif 'Hair Mask Repair Moisturaizing'.lower() in service_name.lower():
        service_name = 'Hair Mask Repair Moisturaizing'
        category = 'treatment'
    elif 'Hair Mask Repairing'.lower() in service_name.lower():
        service_name = 'Hair Mask Repairing'
        category = 'treatment'
    elif 'shampo blessing 1000ml' in service_name.lower():
        service_name = 'Shampo Blessing 1000ml'
        category = 'shampoo'
    elif 'Volumizing shampo 1000 ml'.lower() in service_name.lower():
        service_name = 'Naturica Volumizing Shampo 1000ml'
        category = 'product'
    elif 'Milk Shake Volumizing Shampoo New 300ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Volumizing Shampoo 300ml'
        category = 'product_milk_shake'
    elif 'shampo' in service_name.lower():
        service_name = 'Shampoo'
        category = 'shampoo'
    elif 'Milk Shake Volumizing Shampoo New 300ml'.lower() in service_name.lower():
        service_name = 'Milk Shake Volumizing Shampoo 300ml'
        category = 'product_milk_shake'
    elif 'smoothing l' in service_name.lower():
        service_name = 'Smoothing Long'
        category = 'chemical'
    elif 'dry scalp' in service_name.lower():
        service_name = 'Treatment Dry Scalp'
        category = 'treatment'
    elif 'hairmask color expert' in service_name.lower():
        service_name = 'Hairmask Color Expert'
        category = 'treatment'
    elif 'smooth expert' in service_name.lower():
        service_name = 'Smooth Expert Treatment'
        category = 'treatment'
    elif 'scrub' in service_name.lower():
        service_name = 'Scrub'
        category = 'treatment'
    elif 'dry' in service_name.lower():
        service_name = 'Dry Only'
        category = 'blow'
    elif 'foot polish opie' in service_name.lower():
        service_name = 'Foot Polish OPI'
        category = 'polish'
    elif 'Foot Frecnh Gel Opi'.lower() in service_name.lower():
        service_name = 'Foot French Polish OPI'
        category = 'polish'
    elif 'full under arm rica wax' in service_name.lower():
        service_name = 'Full Under Arm Rica Wax'
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
    elif 'balancing remedy shp 250ml' in service_name.lower():
        service_name = 'Balancing Remedy Shp 250ml'
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
    elif 'energizing miracle shmp 1000ml' in service_name.lower():
        service_name = 'Energizing Miracle Shampoo 1000ml'
        category = 'product'
    elif 'roots spray pemakaian' in service_name.lower():
        service_name = 'Roots Spray Pemakaian'
        category = 'product'
    elif 'volumizing exp condi 1000ml' in service_name.lower():
        service_name = 'Volumizing Exp Conditioner 1000ml'
        category = 'product'
    elif 'soo thing in t relief treatment spray 100ml' in service_name.lower():
        service_name = 'Soothing Relief Treatment Spray 100ml'
        category = 'product'
    elif 'extra loreal 1' in service_name.lower():
        service_name = 'Extra Loreal 1'
        category = 'product'
    elif 'hair coloring short' in service_name.lower():
        service_name = 'Coloring Short'
        category = 'chemical'
    elif 'color hena' in service_name.lower():
        service_name = 'Coloring Hena'
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
    elif 'repair 1/2' in service_name.lower():
        service_name = 'Repair 1/2'
        category = 'extention'
    elif 'pelepasan he' in service_name.lower():
        service_name = 'Lepas Hair Extention'
        category = 'extention'
    elif 'haircut women' in service_name.lower():
        service_name = 'Haircut Woman'
        category = 'haircut'
    elif 'higlight' in service_name.lower():
        service_name = 'Highlight'
        category = 'chemical'
    elif 'repair 1/4' in service_name.lower():
        service_name = 'Repair 1/4'
        category = 'extention'
    elif 'repair 1 tape' in service_name.lower():
        service_name = 'Repair 1 Tape'
        category = 'extention'
    elif 'shape matt' in service_name.lower():
        service_name = 'Shape Matt Putty for Men'
        category = 'product'
    elif 'smoothing poni' in service_name.lower():
        service_name = 'Smoothing Poni'
        category = 'chemical'
    elif 'smoothing s' in service_name.lower():
        service_name = 'Smoothing Short'
        category = 'chemical'
    elif 'color product sendiri' in service_name.lower():
        service_name = 'Color Product Sendiri'
        category = 'chemical'
    elif 'smoothing l' in service_name.lower():
        service_name = 'Smoothing Long'
        category = 'chemical'
    elif 'smothing coco' in service_name.lower():
        service_name = 'Smoothing Coco'
        category = 'chemical'
    elif 'extra smoothing' in service_name.lower():
        service_name = 'Extra Smoothing'
        category = 'product'
    elif 'smoothing' in service_name.lower():
        service_name = 'Smoothing'
        category = 'chemical'
    elif 'perm' in service_name.lower():
        service_name = 'Perming'
        category = 'chemical'
    elif 'makeup' in service_name.lower():
        service_name = 'Makeup'
        category = 'makeup'
    elif 'make up' in service_name.lower():
        service_name = 'Makeup'
        category = 'makeup'
    elif 'make up eye' in service_name.lower():
        service_name = 'Make Up Eye'
        category = 'makeup'
    elif 'eye apply' in service_name.lower():
        service_name = 'Eye Apply'
        category = 'makeup'
    elif 'Eye Shaping'.lower() in service_name.lower():
        service_name = 'Eye Shaping'
        category = 'makeup'
    elif 'Volumizing shampo 1000 ml'.lower() in service_name.lower():
        service_name = 'Naturica Volumizing Shampo 1000ml'
        category = 'product'
    elif 'eyelash' in service_name.lower():
        service_name = 'Eyelash'
        category = 'makeup'

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
            'product_milk_shake': {
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
            'massage': {
                'reduction': 0,
                'commission': 40
            },
            'shampoo': {
                'reduction': 0,
                'commission': 50
            },
            'mp': {
                'reduction': 0,
                'commission': 20
            },
            'nail_art': {
                'reduction': 0,
                'commission': 20
            },
            'nail_extention': {
                'reduction': 0,
                'commission': 30
            },
            'treatment': {
                'reduction': 0,
                'commission': 15
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
                'commission': 35
            },
            'product': {
                'reduction': 0,
                'commission': 10
            },
            'product_milk_shake': {
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
            },
            'nail_art': {
                'reduction': 0,
                'commission': 20
            },
            'nail_extention': {
                'reduction': 0,
                'commission': 30
            }
        }
    }

def Employee():
    return {
        'adit': 'stylist',
        'apriyanti': 'therapies',
        'febriana': 'therapies',
        'marni': 'therapies',
        'angga': 'therapies',
        'narko': 'stylist',
        'jaenal': 'stylist',
        'tama': 'stylist',
        'aryo': 'stylist',
        'wanti': 'stylist',
        'win.d': 'stylist',
        'juna': 'stylist',
        'nuy ': 'therapies',
        'saini': 'therapies',
        'endang ': 'therapies',
        'rohma': 'therapies',
        'cici': 'therapies',
        'tatang': 'stylist',
        'sinta': 'therapies',
        'puput ': 'therapies',
        'nabila ': 'therapies',
        'dilla': 'therapies',
        'fitri': 'therapies',
        'adi ': 'therapies',
        'nia ': 'therapies'
    }
