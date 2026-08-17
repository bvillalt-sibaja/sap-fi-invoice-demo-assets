# SAP FI/AP Invoice Demo Assets

Dependencies for the "Process New Non-PO Invoice Request" .robot (Merz POC):

- `main.py` — SAP GUI FI/AP invoice mirror app
- `excel_mirror.py` — Excel vendor-list workbook mirror app
- `freight_gl_mirror.py` — Excel Freight-tab GL-coding mockup
- `Invoice_WM2201947.pdf` — demo invoice (Waste Mgmt), used by the Acrobat-supplier iteration

Fetched at runtime by the .robot (`Materialize Bundled Dependencies`) so the whole automation
travels as a single .robot file.
