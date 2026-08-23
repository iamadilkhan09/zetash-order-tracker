import streamlit as st
import pandas as pd
import re
import io
from thefuzz import process

st.set_page_config(page_title="Zetash Order Engine", page_icon="📦", layout="wide")
st.title("📦 Zetash & HFZ Automated Order Tracker")
st.markdown("Upload your PostEx CSV to generate a fully formula-driven, live-updating Excel Tracker.")

def normalize_text(text):
    text = str(text).upper()
    noise_words = [
        "REFRENCE", "REFERENCE", "REF", "TALHA", "HASHIR", "ADNAN", "SULAIMAN",
        "ADVERTASMENT", "ADVERTISEMENT", "ADVERTISMENT", "PRODUCTS", "PRODUCT"
    ]
    for nw in noise_words:
        text = re.sub(rf'\b{nw}\b', '', text)
    
    text = re.sub(r'\b(ZETSH|ZETAH|ZETA|ZEETASH|ZETASHH)\b', 'ZETASH', text)
    text = re.sub(r'\b(HFZZ|HZF|HF)\b', 'HFZ', text)
    text = re.sub(r'\b(MUNGRALA|MUNGRELA|MUNGRELLA)\b', 'MUNGRELLA', text)
    
    replacements = {
        r'\b(FAEWASH|FACWASH|FAICEWASH|FACE WASH|FW)\b': 'FACEWASH',
        r'\b(CRAMY|CREMMY|CREAMEY)\b': 'CREAMY',
        r'\b(DAIMAND|DAIOMAND|DIMOND|DIOMAND)\b': 'DIAMOND',
        r'\b(BUEATY|BEAUTI|BEUTY)\b': 'BEAUTY',
        r'\b(BIOTUQUE|BIOTIC|BIOTIQ|BIOTUQE)\b': 'BIOTIQUE',
        r'\b(SCRAB|SCRAP)\b': 'SCRUB',
        r'\b(BLEECH|BLEAK)\b': 'BLEACH',
        r'\b(SUNBLACK|SUN\s*BLOCK|SUN\s*BLACK)\b': 'SUNBLOCK',
        r'\b(SACHEY)\b': 'SACHET',
        r'\b(SOP)\b': 'SOAP',
        r'\b(LOTN)\b': 'LOTION',
        r'\b(PALISH)\b': 'POLISH',
        r'\b(TOOTHPAST)\b': 'TOOTHPASTE',
        r'\b(VASLINE|VASELEN)\b': 'VASELINE',
        r'\b(TUPE|TUB)\b': 'TUBE',
        r'\b(ALVOERA|ALOEVERA|ALOVEERA|ALOVERA)\b': 'ALOE VERA',
        r'\b(ACENE)\b': 'ACNE',
        r'\b(VITAMAN)\b': 'VITAMIN',
        r'\b(GENSIING|GENSENG|GINSENG)\b': 'GINGSENG',
    }
    for pattern, replacement in replacements.items():
        text = re.sub(pattern, replacement, text)
    return re.sub(r'\s+', ' ', text).strip()

# Hardcoded master price list matching EXACT confirmed costs
master_data = [
    ["ZETASH VIP + HAIR COLOR GEL (PUMP)", "Zetash", 1500, 790],
    ["ZETASH VIP + HAIR COLOR GEL (BOTTLE)", "Zetash", 1500, 790],
    ["ZETASH VIP + HAIR COLOR SHAMPOO SACHET", "Zetash", 800, 500],
    ["ZETASH BIOTIQUE SHAMPOO", "Zetash", 600, 375],
    ["ZETASH HAIR PLUS OIL 7 IN 1 FORMULA", "Zetash", 550, 355],
    ["ZETASH HAIR PLUS OIL (ALL VARIANTS)", "Zetash", 500, 270],
    ["ZETASH 3 IN 1 FACEWASH", "Zetash", 450, 175],
    ["ZETASH CREAMY RICE FW", "Zetash", 500, 235],
    ["ZETASH DIAMOND FORMULA CREAM", "Zetash", 400, 290],
    ["ZETASH 24K GOLD FACE SERUM", "Zetash", 999, 450],
    ["ZETASH BEAUTY AGE FACE SERUM", "Zetash", 999, 450],
    ["ZETASH ANTI-ACNE FACE SERUM", "Zetash", 999, 450],
    ["ZETASH BEAUTY CREAM", "Zetash", 300, 220],
    ["ZETASH BEAUTY SERUM", "Zetash", 200, 95],
    ["ZETASH HERBAL 2 IN 1 FORMULA CREAM", "Zetash", 350, 250],
    ["ZETASH BEAUTY SOAP", "Zetash", 200, 90],
    ["ZETASH RICE WATER SKIN TONER", "Zetash", 350, 170],
    ["ZETASH FACIAL SCRUB JAR (ALL VARIANTS)", "Zetash", 500, 240],
    ["ZETASH HAND & FOOT TUBE", "Zetash", 500, 295],
    ["ZETASH 7B'S HAIR SERUM", "Zetash", 0, 0],
    ["ZETASH SOFT COLD CREAM", "Zetash", 400, 160],
    ["ZETASH BODY LOTION (ALL VARIANTS)", "Zetash", 400, 150],
    ["ZETASH SUNBLOCK (SPF 60) SMALL", "Zetash", 500, 220],
    ["ZETASH WELL BLACK MASK 3D FACIAL", "Zetash", 500, 240],
    ["ZETASH HERBAL TOOTHPASTE", "Zetash", 350, 154],
    ["ZETASH DE-HAIR CREAM SACHET (12PC)", "Zetash", 300, 180],
    ["ZETASH STRIP WAX (BOTH VARIANT)", "Zetash", 450, 150],
    ["ZETASH HAIR REMOVER TUBE", "Zetash", 250, 150],
    ["ZETASH HAND & FOOT BLEACH SACHET", "Zetash", 100, 38],
    ["ZETASH CHARCOAL CREAM BLEACH SACHET", "Zetash", 70, 31],
    ["ZETASH CREAM BLEACH JAR", "Zetash", 400, 120],
    ["ZETASH BUST OIL", "Zetash", 800, 380],
    ["ZETASH BIO UP", "Zetash", 700, 320],
    ["ZETASH BIG BUST", "Zetash", 650, 260],
    ["ZETASH HEEL GUARD", "Zetash", 300, 150],
    ["ZETASH VASELINE SMALL", "Zetash", 300, 120],
    ["ZETASH VASELINE LARGE", "Zetash", 400, 220],
    ["ZETASH PRICKLY HEAT SPRAY", "Zetash", 300, 160],
    ["ZETASH GLYCERIN ILARGE", "Zetash", 400, 180],
    ["ZETASH GLYCERIN SMALL", "Zetash", 250, 90],
    ["ZETASH ROSE WATER", "Zetash", 200, 40],
    ["ZETASH ROSE WATER SPRAY", "Zetash", 300, 70],
    ["HFZ ACNE FACE WASH", "HFZ", 450, 175],
    ["HFZ 2 IN 1 VITAMIN C CREAM", "HFZ", 450, 225],
    ["HFZ VITAMIN C FACE WASH", "HFZ", 450, 225],
    ["HFZ CHAWAL FACEWASH", "HFZ", 550, 300],
    ["HFZ VITAMIN C SERUM", "HFZ", 999, 450],
    ["HFZ SWEETY + CREAMY FACEWASH", "HFZ", 550, 390],
    ["HFZ URGENT FACIAL TUBE", "HFZ", 600, 290],
    ["HFZ GINGSENG HAIR OIL", "HFZ", 550, 290],
    ["HFZ GINGSENG KERATIN SHAMPOO", "HFZ", 600, 375],
    ["HFZ MINT BLEACH SACHET", "HFZ", 70, 40],
    ["HFZ PRO CARE BODY LOTION TUBE", "HFZ", 600, 290],
    ["HFZ HAND & FOOT TUBE", "HFZ", 500, 350],
    ["HFZ SCRUB TUBE", "HFZ", 500, 290],
    ["HFZ ACNE TUBE", "HFZ", 300, 180],
    ["HFZ HAIR PLUS OIL", "HFZ", 300, 175],
    ["HFZ SKIN POLISH BOTTLE", "HFZ", 999, 450],
    ["HFZ BREAST ENLARGEMENT JAR", "HFZ", 700, 360],
    ["HFZ URGENT FACIAL SACHET", "HFZ", 50, 40],
    ["MUNGRELLA BLACK SEED OIL", "Mungrella", 680, 530],
    ["MUNGRELLA SPINCER SYRUP", "Mungrella", 390, 320],
    ["MUNGRELLA SCABIES SOAP", "Mungrella", 290, 210]
]
master_products = [x[0] for x in master_data]

aliases = {
    "ZETASH BLACK SEED SHAMPOO": "ZETASH BIOTIQUE SHAMPOO",
    "ZETASH BLACK SEED": "ZETASH BIOTIQUE SHAMPOO",
    "BLACK SEED SHAMPOO": "ZETASH BIOTIQUE SHAMPOO",
    "ZETASH BIOTIQUE ONION SHAMPOO": "ZETASH BIOTIQUE SHAMPOO",
    "ONION SHAMPOO": "ZETASH BIOTIQUE SHAMPOO",
    "7IN1 OIL": "ZETASH HAIR PLUS OIL 7 IN 1 FORMULA",
    "ZETASH 7IN1 OIL": "ZETASH HAIR PLUS OIL 7 IN 1 FORMULA",
    "ZETASH DIAMOND CREAM": "ZETASH DIAMOND FORMULA CREAM",
    "ZETASH DIAMOND": "ZETASH DIAMOND FORMULA CREAM",
    "ZETASH CREAMY FACEWASH": "ZETASH CREAMY RICE FW",
    "ZETASH 3IN1 FACEWASH": "ZETASH 3 IN 1 FACEWASH",
    "HFZ ACNE CLEAR FACEWASH": "HFZ ACNE FACE WASH",
    "HFZ SWEETY FACEWASH": "HFZ SWEETY + CREAMY FACEWASH",
    "ZETASH HERBAL CREAM": "ZETASH HERBAL 2 IN 1 FORMULA CREAM",
    "ZETASH HAND AND FOOT": "ZETASH HAND & FOOT TUBE",
    "ZETASH BEAUTY CREAM": "ZETASH BEAUTY CREAM",
    "HFZ BEAUTY CREAM": "HFZ 2 IN 1 VITAMIN C CREAM",
    "ZETASH SUNBLOCK": "ZETASH SUNBLOCK (SPF 60) SMALL",
    "MUNGRELLA BLACK SEED OIL": "MUNGRELLA BLACK SEED OIL", 
    "SCABIES SOAP": "MUNGRELLA SCABIES SOAP"
}

csv_file = st.file_uploader("Upload PostEx CSV File", type=['csv'])

if csv_file and st.button("🚀 Generate Formula-Driven Tracker", type="primary"):
    with st.spinner("Building live Excel architecture..."):
        raw_df = pd.read_csv(csv_file)
        
        # 1. Prepare Order Items Data
        order_items = []
        for index, row in raw_df.iterrows():
            order_ref = row['ORDER_REFERENCE_NUMBER']
            tracking_no = row['TRACKING_NUMBER']
            raw_detail = str(row['ORDER_DETAIL']).lower()
            
            if pd.isna(row['ORDER_DETAIL']):
                continue
                
            raw_detail = re.sub(r'advance\s*\d*', '', raw_detail) # Strip advance text
            raw_detail = re.sub(r'(?<!\d)\s*pcs\b', ' 1pcs', raw_detail) # Fix naked 'pcs'
            
            parts = re.split(r'(\d+\s*[\W_]*pcs?)', raw_detail)
            current_text = ""
            for part in parts:
                if 'pcs' in part or 'pc' in part:
                    qty_match = re.search(r'\d+', part)
                    qty = int(qty_match.group()) if qty_match else 1
                    clean_text = normalize_text(current_text)
                    if clean_text in aliases:
                        clean_text = aliases[clean_text]
                    matched_product, score = process.extractOne(clean_text, master_products) if clean_text else ("ZETASH BIOTIQUE SHAMPOO", 0)
                    
                    order_items.append({
                        'Order Ref': order_ref,
                        'Tracking # (Link Key)': tracking_no,
                        'Raw Text Segment': current_text.strip().upper(),
                        'Product (dropdown)': matched_product,
                        'Confidence': score,
                        'Qty': qty
                    })
                    current_text = ""
                else:
                    current_text += part
        
        items_df = pd.DataFrame(order_items)
        
        # 2. Prepare Order Tracker Data
        tracker_df = raw_df[['ORDER_REFERENCE_NUMBER', 'TRACKING_NUMBER', 'TRANSACTION_DATE', 'MERCHANT_TRANSACTION_STATUS', 'CUSTOMER_NAME', 'CUSTOMER_PHONE', 'CITY_NAME', 'DELIVERY_ADDRESS', 'ORDER_DETAIL', 'INVOICE_PAYMENT', 'ORDER_PICKUP_DATE', 'ORDER_DELIVERY_DATE', 'REVERSAL_DATE']].copy()
        tracker_df.columns = ['Order Ref', 'Tracking # (Link Key)', 'Order Date', 'Status', 'Customer Name', 'Phone', 'Destination City', 'Delivery Address', 'Original Order Detail (raw)', 'Invoice Amount', 'Pickup Date', 'Delivery Date', 'Reversal Date']
        
        def extract_advance(text):
            detail = str(text).lower()
            match = re.search(r'advance\s*[\W_]*(\d+)', detail)
            return int(match.group(1)) if match else 250 if 'advance' in detail else 0

        tracker_df['Advance Value'] = tracker_df['Original Order Detail (raw)'].apply(extract_advance)
        
        # --- THE FIX: Replace missing dates/empty cells with blank text to prevent Excel errors ---
        tracker_df = tracker_df.fillna("")
        items_df = items_df.fillna("")
        # ----------------------------------------------------------------------------------------------

        # 3. Write native Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # --- SHEET 1: PRICE LIST ---
            price_ws = workbook.add_worksheet('Price List')
            headers = ["Product Name", "Brand", "Selling Price (PKR)", "Cost Price", "Profit per Unit"]
            for col_num, data in enumerate(headers):
                price_ws.write(0, col_num, data)
            for row_num, row_data in enumerate(master_data, 1):
                price_ws.write(row_num, 0, row_data[0])
                price_ws.write(row_num, 1, row_data[1])
                price_ws.write(row_num, 2, row_data[2])
                price_ws.write(row_num, 3, row_data[3]) # Write confirmed real Cost Price
                price_ws.write_formula(row_num, 4, f'=C{row_num+1}-D{row_num+1}') # Profit calculation
            
            # --- SHEET 2: ORDER TRACKER ---
            tracker_ws = workbook.add_worksheet('Order Tracker')
            tracker_ws.write('P1', 'Delivery Charge')
            tracker_ws.write('P2', 250)
            tracker_ws.write('Q1', 'PostEx Charge')
            tracker_ws.write('Q2', 225)
            
            t_headers = tracker_df.columns.tolist() + ['Total Selling Price', 'Total Cost', 'Estimated Profit']
            for col_num, data in enumerate(t_headers):
                tracker_ws.write(0, col_num, data)
                
            for row_num, row_data in enumerate(tracker_df.values, 1):
                for col_num, data in enumerate(row_data):
                    tracker_ws.write(row_num, col_num, data)
                
                # Formulas
                tracking_cell = f'B{row_num+1}'
                inv_cell = f'J{row_num+1}'
                adv_cell = f'N{row_num+1}'
                status_cell = f'D{row_num+1}'
                
                tracker_ws.write_formula(row_num, 14, f"=SUMIF('Order Items'!B:B, {tracking_cell}, 'Order Items'!F:F)")
                tracker_ws.write_formula(row_num, 15, f"=SUMIF('Order Items'!B:B, {tracking_cell}, 'Order Items'!H:H)")
                tracker_ws.write_formula(row_num, 16, f'=IF(OR({status_cell}<>"Delivered", {inv_cell}=0), 0, ({inv_cell}+{adv_cell})-P{row_num+1}-$Q$2)')
            
            # Dropdowns & Formatting for Tracker
            status_list = ['Delivered', 'Return', 'In Transit', 'Pending', 'Under Review', 'Unbooked', 'Attempted', 'Cancelled']
            tracker_ws.data_validation(f'D2:D{len(tracker_df)+1}', {'validate': 'list', 'source': status_list})
            
            red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            light_green_fmt = workbook.add_format({'bg_color': '#D8E4BC', 'font_color': '#000000'})
            
            tracker_ws.conditional_format(f'D2:D{len(tracker_df)+1}', {'type': 'cell', 'criteria': '==', 'value': '"Return"', 'format': red_fmt})
            tracker_ws.conditional_format(f'D2:D{len(tracker_df)+1}', {'type': 'cell', 'criteria': '==', 'value': '"Delivered"', 'format': green_fmt})
            
            # Format Invoice Amount column with light green if it equals 0
            tracker_ws.conditional_format(f'J2:J{len(tracker_df)+1}', {'type': 'cell', 'criteria': '==', 'value': 0, 'format': light_green_fmt})

            # --- SHEET 3: ORDER ITEMS ---
            items_ws = workbook.add_worksheet('Order Items')
            i_headers = ['Order Ref', 'Tracking #', 'Raw Text Segment', 'Product (dropdown)', 'Qty', 'Selling Price', 'Line Selling Total', 'Line Cost', 'Line Profit']
            for col_num, data in enumerate(i_headers):
                items_ws.write(0, col_num, data)
                
            for row_num, row_data in enumerate(items_df.values, 1):
                items_ws.write(row_num, 0, row_data[0])
                items_ws.write(row_num, 1, row_data[1])
                items_ws.write(row_num, 2, row_data[2])
                items_ws.write(row_num, 3, row_data[3])
                items_ws.write(row_num, 4, row_data[5])
                
                prod_cell = f'D{row_num+1}'
                qty_cell = f'E{row_num+1}'
                
                items_ws.write_formula(row_num, 5, f"=IFERROR(VLOOKUP({prod_cell}, 'Price List'!A:E, 3, FALSE), 0)")
                items_ws.write_formula(row_num, 6, f"={qty_cell}*F{row_num+1}")
                items_ws.write_formula(row_num, 7, f"=IFERROR(VLOOKUP({prod_cell}, 'Price List'!A:E, 4, FALSE)*{qty_cell}, 0)")
                items_ws.write_formula(row_num, 8, f"=G{row_num+1}-H{row_num+1}")
                
                # Flag bad matches (Confidence < 85)
                if row_data[4] < 85:
                    items_ws.write(row_num, 3, row_data[3], workbook.add_format({'bg_color': '#FFEB9C'}))
            
            # Product Dropdown
            items_ws.data_validation(f'D2:D{len(items_df)+1}', {'validate': 'list', 'source': f"='Price List'!$A$2:$A${len(master_data)+1}"})

            # --- SHEET 4: SUMMARY ---
            sum_ws = workbook.add_worksheet('Summary')
            sum_ws.write('A1', 'Total Orders')
            sum_ws.write_formula('B1', f"=COUNTA('Order Tracker'!A2:A{len(tracker_df)+1})")
            sum_ws.write('A2', 'Delivered Orders')
            sum_ws.write_formula('B2', f"=COUNTIF('Order Tracker'!D2:D{len(tracker_df)+1}, \"Delivered\")")
            sum_ws.write('A3', 'Total Estimated Profit')
            sum_ws.write_formula('B3', f"=SUM('Order Tracker'!Q2:Q{len(tracker_df)+1})")
            
        st.success("✅ Fully Interactive Excel Architecture Built!")
        st.download_button(label="📥 Download Live Excel Tracker", data=output.getvalue(), file_name="Zetash_Live_Tracker.xlsx")
