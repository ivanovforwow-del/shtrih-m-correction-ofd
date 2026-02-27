import sys
sys.path.insert(0, '.')
from parse_ofd_receipts_final import fetch_receipt_data, extract_receipt_params
import re
import html as html_module

# Test with a known receipt hyperlink that has KT EAN-13
# From receipts_data.csv: fp=1852047038 should have one
hyperlink = "https://lk.platformaofd.ru/web/noauth/cheque?id=155485775598&date=1767201399000&fp=1852047038"

params = extract_receipt_params(hyperlink)
print(f"Params: {params}")

html_content = fetch_receipt_data(params)
if html_content:
    print(f"HTML content length: {len(html_content)}")
    
    # Find receipt container
    match = re.search(r'id="fido_cheque_container">(.*?)</div>\s*<div', html_content, re.DOTALL)
    if not match:
        match = re.search(r'id="fido_cheque_container">(.*?)</div>$', html_content, re.DOTALL)
    
    if match:
        decoded_html = html_module.unescape(match.group(1))
        print(f"Decoded HTML length: {len(decoded_html)}")
        
        # Find item names
        item_names = re.findall(r'<b>(\d+):\s*([^<]+)</b>', decoded_html)
        print(f"Found {len(item_names)} items")
        
        # Find marking codes
        # Let's try different patterns
        print("\n=== Trying different patterns ===")
        
        # Pattern 1: KT EAN-13
        pattern1 = r'КТ EAN-13.*?<span>([\d]+)</span>'
        matches1 = re.findall(pattern1, decoded_html)
        print(f"Pattern 1 (КТ EAN-13): {matches1}")
        
        # Pattern 2: with DOTALL
        pattern2 = r'КТ EAN-13.*?<span>([\d]+)</span>'
        matches2 = re.findall(pattern2, decoded_html, re.DOTALL)
        print(f"Pattern 2 with DOTALL: {matches2}")
        
        # Let's look at what's around the marking code
        if 'КТ EAN-13' in decoded_html:
            idx = decoded_html.index('КТ EAN-13')
            print(f"\nContext around KT EAN-13:")
            print(repr(decoded_html[idx:idx+200]))
    else:
        print("No fido_cheque_container found!")
else:
    print("No HTML content returned")