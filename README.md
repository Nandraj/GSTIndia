# GST (India) Tools

[1. Statement-1A Excel to Json Generator](#statement-1a-excel-to-json-generator) [ [ codes ](Stmt_1A_Json_Generator/main.py) ]

## Statement-1A Excel to Json Generator
We faced many problems with government provided excel macro based utility especially with regards to date columns. Generally we prepare this statement in regular excel file for printing and signature. And as it contains many rows obviously we want to directly copy this data into excel utility but date format get changed for all rows and excel utility is complaining for that. 

Also utility only work with computers having Microsoft Office 2010 or later but some user may have 2007 version.

So I developed [this](Stmt_1A_Json_Generator/main.py) utility to make my work less painful overnight. Although I have to update codes to cope with different date format but after that it works just fine. Just put GSTIN, tax period and paste all rows into given excel template and on single click json is generated at the place where excel file is located.