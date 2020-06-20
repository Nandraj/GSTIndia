# GST (India) Tools

[1. Statement-1A Excel to Json Generator](#statement-1a-excel-to-json-generator) [ [ source ](Stmt_1A_Json_Generator/main.py) ] [ [ download ](https://app.box.com/s/s03hp5lzzobiock7p1bdmxogutbzxkw5)]

[2. GSTR-2A Excel Formatter](#gstr-2a-excel-formatter) [ [ source ](GSTR_2A_Excel_Formatter/main.py) ] [ [ download ](https://app.box.com/s/k5eljcv6h7q66x6m8mwbojqjc3vt47yc)]

## Statement-1A Excel to Json Generator
We faced many problems with government provided excel macro based utility especially with regards to date columns. Generally we prepare this statement in regular excel file for printing and signature. And as it contains many rows obviously we want to directly copy this data into excel utility but date format get changed for all rows and excel utility is complaining for that. 

Also utility only work with computers having Microsoft Office 2010 or later but some user may have 2007 version.

So I developed [this](Stmt_1A_Json_Generator/main.py) utility to make my work less painful overnight. Although I have to update codes to cope with different date format but after that it works just fine. Just put GSTIN, tax period and paste all rows into given excel template and on single click json is generated at the place where excel file is located.

## GSTR-2A Excel Formatter
Sometimes we need to prepare GSTR-2A (for refund purpose to get it printed) for multiple months but government website (gst.gov.in) only provides facility to download files for individual months. Excel file for a year can be generated from other proprietory software but it is not in govt. provided format so I developed this tool which takes that other software generated excel and reformat it into govt. format then only manual work remains is of pasting govt. 2A formatted headers with logo.