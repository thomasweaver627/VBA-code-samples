# VBA-code-samples


## ReportExport()
VBA function that takes a file name string, directory string, and SELECT statement string as arguments, and exports the records to an .xlsx file with some basic formatting.

Using an array of "A" to "Z", it is designed to handle any number of queried fields. And while typically unnecessary, it can also be modified to include "AA", "AB", "AC", etc. as needed.

It can also be easily modified to accept a stored Access query instead, using DoCmd.TransferSpreadsheet.

    *Note: current copy is the first version; further changes/testing exposed some errors*
