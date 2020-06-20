# VBA-Excel-Array2_Filter
VBA Excel Filter a two-dimensional array by different columns with a variety of

# How to use?
The A2_Filter_AND function waits for two arguments:
- a two-dimensional array with data to be filtered
- a two-dimensional array of criteria.
 
 Array of criteria = two-dimensional array.
in column 1 - column numbers
in the 2nd column - filter value (row, number, date)
in column 3 - filtration method (with a sign of case accounting)
 
 The methods of filtering are implemented in the function Criteria_Check.
"EQUAL_TEXT" ' the lines are equal without case.
"EQUAL_BINA" ' the lines are equal case sensitive.
"CONTA_TEXT" ' the line contains WITHOUT case accounting
"CONTA_BINA" ' the string contains case sensitive
 
 An example of a criteria string:
 9, "Budget", "CONTA_BINA" - leave the rows that contain "Budget" in the ninth column, case sensitive
 
Function A2_Filter_AND Returns the array two-dimensional filtered.

# Tested 
[![Rubberduck](//https://user-images.githubusercontent.com/5751684/48656196-a507af80-e9ef-11e8-9c09-1ce3c619c019.png/150x100)](https://github.com/rubberduck-vba/Rubberduck/)

# Plans
Implement date and number filters.
