# What this VBA does
It is for daily purchase. It indicates which orders are bought already, and updates them in the <Purchased Parts (Pad + Rotor).xlsx> excel workbook automatically. If not bought, then generates a new workbook called <qbp.csv> and a list to search them easily on other websites. 

# What this file contains
Daily PO VBA: It is the VBA to run.\
po_wo_20000101.xlsx: It is the main file to run the VBA in.\
Purchased Parts (Pad + Rotor).xlsx, po_wo_20000101-0715S.xlsx, and pads_hardware.xlsx: They need to be downloaded in the same folder, and may remain closed before running the VBA.

# How it works
1. Make sure all the excel files are downloaded in the same folder, and make sure the file_dir in the VBA has changed correspondingly. 
2. Open <po_wo_20010101.xlsx>.
3. Run sub "Daily_1", then search orders on QBP using the newly generated <qbp.csv>.
4. Run sub "Daily_2" to generate the list needed for Transbec search.
5. Run sub "Daily_3" to copy & paste data to workbook <po_wo_20010101-0715S.xlsx> and color them.
