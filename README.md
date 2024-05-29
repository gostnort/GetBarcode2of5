# GetBarcode2of5

This Microsoft Excel VBA module has a function named 'GetBarcode2of5()' and a substance named 'ConvertSelected2Barcode'. 

## Import to Excel

There is a file of 'Decimal2Barcode_Module.bas'. You can download this file; access to the 'Visual Basic' in Excel. Right click the project tree window, and select 'Import File'.

You will see the 'ConvertSelected2Barcode' displayed in your Marco. There is a function in any cells after you typed '=GetBarcode2of5'.

## Usage

Only numbers are supported for **2of5 Interleaved** barcode generation. The barcode will draw on the cell which has the number. **Do not process over** the limit of the type of ***Currency*** in VBA. That's about **15-digits**.

### ConvertSelected2Barcode

Running this marco will add the relevant barcode on the selected cell(s).

### GetBarcode2of5

This function will add the relevant barcode on the cell. To provent endless looping the event of "Worksheet_Change", the content of cell will be changed to the barcode number without this function calling.


