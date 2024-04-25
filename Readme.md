# newSSExcel Addin

## Instructions

TBD

This is the master example for authentication / single sign on

git clone --bare https://github.com/MooreAzata/Space_Station_Excel_Addin.git

git push --mirror https://github.com/MooreAzata/newSSExcel.git   

# EXCEL SHEET MARKERS

#C1B, #C2B ... #C12B -
Put above columns containing budget data, with the number pointing to the month index (#C1B for January, #C2B for February etc.)

#C1A, #C2A ... #C12A -
Put above columns containing actuals data, as above

#R -
Put aside rows containing budget/actuals data. Used for clearing

#RX -
Replaces #R, signifying the final row. The system should work without this, but will run a little slower

#RC -
Put above the column containing the #R and #RX markers

#N -
Put above the column containing RPG codes. Used for writing

#BC -
Put above the column of the cell where the budget name is displayed

#BR -
Put aside the row of the cell where the budget name is displayed