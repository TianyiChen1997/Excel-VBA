# Excel VBA Overview

Currently Available:
<Weekly_Rotor/Pad> for update, cleanup, and calculation for the Weekly Resotck Summary
<Daily_PO> for daily purchase

Working on: 

# Weekly_Rotor/Pad

It is for Rotor/Pad Weekly Restock Summary. It updates the "Total Amount" data and "PowerBI Sales" data, cleans up the worksheets, and calculates the "Buy Qty".  There are two subs, <Rotor/Pad Cleanup> and <Rotor/Pad Buy>. One is for cleanup and data update, and the other one is for calculation. They are called together under the sub <Rotor/Pad Weekly>.

# Daily_PO

It is for daily purchase. It indicates which orders are bought already, and updates them in the "Purchased Parts" excel workbook automatically. If not bought, then generates a new workbook and a list to search them easily on other websites. Three subs need to be run. Run the <Daily_1> sub at the very beginning, then search orders on QBP using the newly generated qbp.csv, and then run the <Daily_2> sub to generate the list needed for Transbec search, and lastly run the <Daily_3> sub to copy & paste data to the original workbook and color them.
