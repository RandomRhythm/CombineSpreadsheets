# CombineSpreadsheets
Combine all columns from two Microsoft Excel spreadsheets into one based on matching column values. If you have two spreadsheets with different data but one data column is the same, then this script can combine those two spreadsheets based on matching values. Example:
Spreadsheet 1 has columns SHA256, File Path
Spreadsheet 2 has columns Threat Score, SHA256
Combined spreadsheet has columns SHA256, File Path, Threat Score, SHA256

Combining off the SHA256 column will add all columns from the first match row in spreadsheet 2 to spreadsheet 1.
