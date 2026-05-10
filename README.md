# Arte Di Casa Billing System — Updated Fixes

## Latest fixes
- Seller VAT / TRN is now printed clearly on invoice/PDF/Word/Excel outputs.
- Product entry now has a full-row delete button.
- Product serial numbers automatically move up after deleting any row.
- No retyping is needed when deleting row 4 from a long product list.
- Packing list split rows now insert immediately below the selected item instead of at the bottom.
- If a product is deleted, its packing rows are removed and packing box numbers reorder automatically.

## Login
Password: 1985

## Run
pip install -r requirements.txt
streamlit run app.py

## Streamlit Cloud
Upload all files to GitHub and reboot.
Main file: app.py

## Word Import Feature
- Go to Create / Edit.
- Upload an extracted Word file (.docx) in Import Extracted Word File.
- Click Import Word.
- Review/edit products and customer details.
- Save as proforma or convert to invoice, then add packing details.

## Word Invoice Import Update
- Invoice DOCX imports now read the packing list table when present.
- Imported packing rows include Box No, Part, Brand, Product Details, L, B, H, CBM, GW and NW.
- Packing summary is rebuilt automatically after import.
