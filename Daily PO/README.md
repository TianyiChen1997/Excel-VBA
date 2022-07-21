# What this VBA does
It is for daily purchase. It indicates which orders are bought already, and updates them in the <Purchased Parts (Pad + Rotor).xlsx> excel workbook automatically. If not bought, then generates a new workbook called <qbp.csv> and a list to search them easily on other websites. 

# What this folder contains
* Daily PO VBA : It is the VBA to run.
* po_wo_20000101.xlsx : It is the main file to run the VBA in.
* All other excel files : They need to be downloaded in the same folder, and may remain closed before running the VBA.

# How it works
1. Make sure all the excel files are downloaded in the same folder, and make sure the file_dir in the VBA has changed correspondingly. 
2. Open <po_wo_20010101.xlsx>.
3. Run sub "Daily_1", then search orders on QBP using the newly generated <qbp.csv>. Enter "QBP" in column E for each item found on the QBP website.
4. Run sub "Daily_2" to generate the list needed for Transbec search. Use the list in column G to J to search items on Transbec, and enter "Transbec" in column E if the item is found. If not found on QBP nor on Transbec, then enter "N/A".
5. Run sub "Daily_3" to copy & paste data to workbook <po_wo_20010101-0715S.xlsx> and color them. You may now open <po_wo_20010101-0715S.xlsx> to double check if the information is entered correctly.
