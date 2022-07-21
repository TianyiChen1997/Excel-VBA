# What this VBA does
There are two VBA here, one for Rotor Weekly Restock Summary and the other for Pad Weekly Restock Summary. They first update the "Total Amount" data using <ta.csv> and "PowerBI Sales" data using <onz.csv>, <cnz.csv>, <afn.csv>, and <cdm.csv>. They then clean up the worksheets, and calculate the "Buy Qty". 

# What this folder contains
* Weekly Pad VBA and Weekly Rotor VBA : They are the VBA to run.
* PAD Summary.xlsm and Rotor Summary.xlsm : They are the main files for Weekly Pad VBA and Weekly Rotor VBA to run respectively.
* All other files : They need to be downloaded in the same folder, and may remain closed before running the VBA.

# How it works
1. Make sure all the files are downloaded in the same folder, and make sure the file_dir in the VBA has changed correspondingly. 
2. Open <PAD Summary.xlsm> or <Rotor Summary.xlsm>.
3. Run sub "Pad_Weekly" or "Rotor_Weekly" based on the opened excel file, and now both sales and total amount data are updated and the quantity to buy is calculated automatically.
