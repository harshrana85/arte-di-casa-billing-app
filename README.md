# Arte Di Casa Billing System — CBM Word Export/Import Fix

## Final Fix
- Manual CBM values in the Packing List now stay saved.
- Word/PDF/Excel exports no longer recalculate manual CBM back to zero.
- Imported Word invoices bring CBM values back into the editable packing list.
- Older Word files that printed row CBM as zero but had a Total CBM summary will preserve the printed total and place it into the packing data instead of losing it.

## Login
Password: 1985

## Run
pip install -r requirements.txt
streamlit run app.py
