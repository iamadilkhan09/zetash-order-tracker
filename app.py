import streamlit as st
import pandas as pd
import re
import numpy as np
from thefuzz import process
import io

# Page Config
st.set_page_config(page_title="Zetash Order Engine", page_icon="📦", layout="wide")

st.title("📦 Zetash & HFZ Automated Order Tracker")
st.markdown("Upload your monthly PostEx CSV export and your master Price List Excel file to instantly generate your financial tracker.")

# --- THE AUTO-CORRECT ENGINE ---
def normalize_text(text):
    text = str(text).upper()
    
    # 1. Remove Staff Noise & References
    noise_words = [
        "REFRENCE", "REFERENCE", "REF", "TALHA", "HASHIR", "ADNAN", 
        "ADVERTASMENT", "ADVERTISEMENT", "PRODUCTS", "PRODUCT"
    ]
    for nw in noise_words:
        text = re.sub(rf'\b{nw}\b', '', text)
    
    # 2. Fix Brand Names
    text = re.sub(r'\b(ZETSH|ZETAH|ZETA|ZEETASH|ZETASHH)\b', 'ZETASH', text)
    text = re.sub(r'\b(HFZZ|HZF|HF)\b', 'HFZ', text)
    text = re.sub(r'\b(MUNGRALA|MUNGRELA|MUNGRELLA)\b', 'MUNGRELLA', text)
    
    # 3. Fix Phonetic Typos & Slang
    replacements = {
        r'\b(FAEWASH|FACWASH|FAICEWASH|FACE WASH|FW)\b': 'FACEWASH',
        r'\b(CRAMY|CREMMY|CREAMEY)\b': 'CREAMY',
        r'\b(DAIMAND|DAIOMAND|DIMOND)\b': 'DIAMOND',
        r'\b(BUEATY|BEAUTI|BEUTY)\b': 'BEAUTY',
        r'\b(BIOTUQUE|BIOTIC|BIOTIQ)\b': 'BIOTIQUE',
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
    }
    
    for pattern, replacement in replacements.items():
        text = re.sub(pattern, replacement, text)
        
    text = re.sub(r'\s+', ' ', text).strip()
    return text

# Drag and Drop Uploader
col1, col2 = st.columns(2)
with col1:
    csv_file = st.file_uploader("1. Upload PostEx CSV File", type=['csv'])
with col2:
    excel_file = st.file_uploader("2. Upload Master Price List (Excel)", type=['xlsx'])

if csv_file and excel_file:
    if st.button("🚀 Generate Financial Tracker", type="primary"):
        with st.spinner("Processing orders, applying auto-correct, and calculating profits..."):
            
            raw_df = pd.read_csv(csv_file)
            price_list_df = pd.read_excel(excel_file, sheet_name='Price List')
            price_list_df = price_list_df.dropna(subset=['Product Name'])
            
            master_products = price_list_df['Product Name'].tolist()
            price_lookup = price_list_df.set_index('Product Name').to_dict(orient='index')
            
            DELIVERY_CHARGE = 250
            POSTEX_CHARGE = 225
            DEFAULT_ADVANCE_VALUE = 250
            
            aliases = {
                "ZETASH BLACK SEED SHAMPOO": "ZETASH BIOTIQUE SHAMPOO",
                "ZETASH BLACK SEED": "ZETASH BIOTIQUE SHAMPOO",
                "BLACK SEED SHAMPOO": "ZETASH BIOTIQUE SHAMPOO",
                "ZETASH BIOTIQUE ONION SHAMPOO": "ZETASH BIOTIQUE SHAMPOO",
                "ONION SHAMPOO": "ZETASH BIOTIQUE SHAMPOO",
                "ZETASH BIOTIQUE BLACK SEED SHAMPOO": "ZETASH BIOTIQUE SHAMPOO",
                "ZETASH BIOTIQUE GREEN SHAMPOO": "ZETASH BIOTIQUE SHAMPOO",
                "ZETASH GREEN SHAMPOO": "ZETASH BIOTIQUE SHAMPOO",
                
                "7IN1 OIL": "ZETASH HAIR PLUS OIL 7 IN 1 FORMULA",
                "ZETASH 7IN1 OIL": "ZETASH HAIR PLUS OIL 7 IN 1 FORMULA",
                "ZETASH ZETASH 7IN1 OIL": "ZETASH HAIR PLUS OIL 7 IN 1 FORMULA",
                
                "ZETASH DIAMOND CREAM": "ZETASH DIAMOND FORMULA CREAM",
                "ZETASH DIAMOND": "ZETASH DIAMOND FORMULA CREAM",
                
                "ZETASH CREAMY FACEWASH": "ZETASH CREAMY RICE FW",
                "ZETASH 3IN1 FACEWASH": "ZETASH 3 IN 1 FACEWASH",
                "HFZ ACNE CLEAR FACEWASH": "HFZ ACNE FACE WASH",
                "HFZ SWEETY FACEWASH": "HFZ SWEETY + CREAMY FACEWASH",
                
                "ZETASH HERBAL CREAM": "ZETASH HERBAL 2 IN 1 FORMULA CREAM",
                "ZETASH HAND AND FOOT": "ZETASH HAND & FOOT TUBE",
                "ZETASH HAND AND FOOT CREAM": "ZETASH HAND & FOOT TUBE",
                "ZETASH BEAUTY CREAM": "ZETASH BEAUTY CREAM",
                "HFZ BEAUTY CREAM": "HFZ 2 IN 1 VITAMIN C CREAM",
                "ZETASH SUNBLOCK": "ZETASH SUNBLOCK (SPF 60) SMALL",
                "ZETASH BLACK MASK TUBE": "ZETASH WELL BLACK MASK 3D FACIAL",
                
                "MUNGRELLA BLACK SEED OIL": "MUNGRELLA BLACK SEED OIL", 
                "BLACK SEED OIL": "MUNGRELLA BLACK SEED OIL",    
                "MUNGRELLA SYRUP": "MUNGRELLA SPINCER SYRUP",
                "SCABIES SOAP": "MUNGRELLA SCABIES SOAP"
            }
            
            order_items = []
            for index, row in raw_df.iterrows():
                order_ref = row['ORDER_REFERENCE_NUMBER']
                tracking_no = row['TRACKING_NUMBER']
                raw_detail = str(row['ORDER_DETAIL']).lower()
                
                if pd.isna(row['ORDER_DETAIL']):
                    continue
                    
                parts = re.split(r'(\d+\s*[\W_]*pcs)', raw_detail)
                current_text = ""
                found_pcs = False
                
                for part in parts:
                    if 'pcs' in part:
                        found_pcs = True
                        qty_match = re.search(r'\d+', part)
                        qty = int(qty_match.group()) if qty_match else 1
                        
                        clean_text = normalize_text(current_text)
                        if clean_text in aliases:
                            clean_text = aliases[clean_text]
                        
                        matched_product, score = process.extractOne(clean_text, master_products)
                        prod_info = price_lookup.get(matched_product, {})
                        
                        order_items.append({
                            'Order Ref': order_ref,
                            'Tracking # (Link Key)': tracking_no,
                            'Brand': prod_info.get('Brand', 'Zetash'),
                            'Product (dropdown)': matched_product,
                            'Qty': qty,
                            'Selling Price (unit)': prod_info.get('Selling Price (PKR)', 0),
                            'Line Selling Total': prod_info.get('Selling Price (PKR)', 0) * qty,
                            'Line Cost Total': prod_info.get('Cost Price', 0) * qty,
                            'Line Profit': (prod_info.get('Selling Price (PKR)', 0) * qty) - (prod_info.get('Cost Price', 0) * qty),
                            'Match Confidence': 'High' if score > 75 else 'Review Needed',
                            'Raw Text Segment': current_text.strip().upper()
                        })
                        current_text = ""
                    else:
                        current_text += part
                        
                if not found_pcs and raw_detail.strip():
                    clean_text = normalize_text(raw_detail)
                    if clean_text in aliases:
                        clean_text = aliases[clean_text]
                        
                    matched_product, score = process.extractOne(clean_text, master_products)
                    prod_info = price_lookup.get(matched_product, {})
                    
                    order_items.append({
                        'Order Ref': order_ref,
                        'Tracking # (Link Key)': tracking_no,
                        'Brand': prod_info.get('Brand', 'Zetash'),
                        'Product (dropdown)': matched_product,
                        'Qty': 1, 
                        'Selling Price (unit)': prod_info.get('Selling Price (PKR)', 0),
                        'Line Selling Total': prod_info.get('Selling Price (PKR)', 0) * 1,
                        'Line Cost Total': prod_info.get('Cost Price', 0) * 1,
                        'Line Profit': (prod_info.get('Selling Price (PKR)', 0) * 1) - (prod_info.get('Cost Price', 0) * 1),
                        'Match Confidence': 'High' if score > 75 else 'Review Needed',
                        'Raw Text Segment': raw_detail.strip().upper()
                    })
                        
            items_df = pd.DataFrame(order_items, columns=['Order Ref', 'Tracking # (Link Key)', 'Brand', 'Product (dropdown)', 'Qty', 'Selling Price (unit)', 'Line Selling Total', 'Line Cost Total', 'Line Profit', 'Match Confidence', 'Raw Text Segment'])
            
            tracker_df = raw_df[['ORDER_REFERENCE_NUMBER', 'TRACKING_NUMBER', 'TRANSACTION_DATE', 'MERCHANT_TRANSACTION_STATUS', 'CUSTOMER_NAME', 'CUSTOMER_PHONE', 'CITY_NAME', 'DELIVERY_ADDRESS', 'ORDER_DETAIL', 'INVOICE_PAYMENT', 'ORDER_PICKUP_DATE', 'ORDER_DELIVERY_DATE', 'REVERSAL_DATE']].copy()
            tracker_df.columns = ['Order Ref', 'Tracking # (Link Key)', 'Order Date', 'Status', 'Customer Name', 'Phone', 'Destination City', 'Delivery Address', 'Original Order Detail (raw)', 'Invoice Amount', 'Pickup Date', 'Delivery Date', 'Reversal Date']
            
            tracker_df['Delivery Chg (Customer)'] = DELIVERY_CHARGE
            tracker_df['PostEx Charge (Us)'] = POSTEX_CHARGE
            
            def extract_advance(text):
                detail = str(text).lower()
                if 'advance' in detail:
                    match = re.search(r'advance\s*[\W_]*(\d+)', detail)
                    if match:
                        return int(match.group(1))
                    return DEFAULT_ADVANCE_VALUE
                return 0

            tracker_df['Advance Value'] = tracker_df['Original Order Detail (raw)'].apply(extract_advance)
            tracker_df['Advance Payment?'] = np.where(tracker_df['Advance Value'] > 0, 'Yes', 'No')
            
            item_totals = items_df.groupby('Tracking # (Link Key)').agg({'Line Selling Total': 'sum', 'Line Cost Total': 'sum'}).reset_index()
            item_totals.columns = ['Tracking # (Link Key)', 'Total Selling Price (from Items)', 'Total Cost (from Items)']
            
            tracker_df = pd.merge(tracker_df, item_totals, on='Tracking # (Link Key)', how='left')
            tracker_df['Total Selling Price (from Items)'] = tracker_df['Total Selling Price (from Items)'].fillna(0)
            tracker_df['Total Cost (from Items)'] = tracker_df['Total Cost (from Items)'].fillna(0)
            
            tracker_df['Estimated Profit'] = np.where(tracker_df['Status'] == 'Delivered', (tracker_df['Invoice Amount'] + tracker_df['Advance Value']) - tracker_df['Total Cost (from Items)'] - tracker_df['PostEx Charge (Us)'], 0)
            
            def generate_flags(row):
                flags = []
                if row['Advance Value'] > 0 and row['Advance Value'] != DEFAULT_ADVANCE_VALUE:
                    flags.append(f"Custom Advance ({row['Advance Value']})")
                if row['Status'] == 'Delivered' and row['Estimated Profit'] < 0:
                    flags.append("Negative Profit")
                return " | ".join(flags) if flags else "OK"

            tracker_df['Needs Manual Review?'] = tracker_df.apply(generate_flags, axis=1)
            
            cols = tracker_df.columns.tolist()
            cols.insert(4, cols.pop(cols.index('Needs Manual Review?')))
            tracker_df = tracker_df[cols]
            
            total_orders = len(tracker_df)
            delivered_orders = len(tracker_df[tracker_df['Status'] == 'Delivered'])
            total_profit = tracker_df['Estimated Profit'].sum()
            
            summary_df = pd.DataFrame({'Metric': ['Total Orders', 'Delivered Orders', 'Total Estimated Profit (PKR)'], 'Value': [total_orders, delivered_orders, total_profit]})
            
            # Display metrics on screen
            st.success("✅ Tracker generated successfully!")
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Orders", total_orders)
            m2.metric("Delivered Orders", delivered_orders)
            m3.metric("Total Estimated Profit", f"PKR {total_profit:,.2f}")
            
            # Download Button
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                price_list_df.to_excel(writer, sheet_name='Price List', index=False)
                tracker_df.to_excel(writer, sheet_name='Order Tracker', index=False)
                items_df.to_excel(writer, sheet_name='Order Items', index=False)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            st.download_button(
                label="📥 Download Final Excel Tracker",
                data=output.getvalue(),
                file_name="Automated_Order_Tracker_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )