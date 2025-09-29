# Automatic Export Excel Invoice to PDF File

This is an example of applying VBA code to create an automatic PDF of an Excel file with just one click.

## File Structure

The Excel file includes 3 tabs: **Invoice**, **Customers**, **Products**.

## Dynamic Selection Setup

In order to create dynamic selection of invoices to send to different companies, a drop-down list using **Data Validation** has been created to shift between companies.

The other company's information such as address and email is automatically updated by using `XLOOKUP()`.

**Example:**
```excel
=XLOOKUP($D$5,Customers!B:B,Customers!C:C)
```

Also, a drop-down list using **Data Validation** is created in the Product Description area to select from 30 different products. The price is automatically updated by using `XLOOKUP()`. At the same time, Tax and Total (exc. Tax) are updated by using calculation formulas.

**Example:**
Unit price: searching the product name in the Invoice tab and returning the result from the Products tab
```excel
=XLOOKUP(B12,Products!C:C,Products!D:D)
```

The calculation formula is still present in empty cells by using the `IF()` function. Only when there is a value in the Unit Price column will it return a value. Otherwise, it will return an empty string.

**Example:**
In Unit Price column:
```excel
=IF(B15="","",XLOOKUP(B15,Products!C:C,Products!D:D))
```
This will only return a unit price if there is a value in Product Description (B15); otherwise, it returns an empty string.

## Date Functions

Invoice Date can use the current date with the function `TODAY()`. Due Date can add extra days such as `TODAY() + 28`.

## Print Area Setup

Set Print Area in the Page Layout menu to ensure Excel will print only the area we selected.

## VBA Automation to Transform File to PDF

```vba
Sub PDF()
    Dim invoice_number As Long
    Dim name As String
    Dim file_path As String
    Dim file_name As String
    
    invoice_number = Range("G4")
    name = Range("D5")
    file_path = "C:\Data Analysis Career Path\Excel\VBA practice\PDF export with VBA\"
    file_name = invoice_number & "_" & name
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, ignoreprintareas:=False, Filename:=file_path & file_name
End Sub
```

Some dimensions such as `invoice_number`, `name`, `file_path`, and `file_name` are created. The file name is created with the combination of `invoice_number` and `name`. Then it is saved to the current local path.

## Creating the Button

To create a button to run this VBA code, in the Developer section, insert a button and change it to "PDF" name. When clicking this button, the created PDF file will be saved to the local path set up before.

**NOTE:** Save the file as **Excel Macro-Enabled Workbook (*.xlsm)** in order to run any VBA macros.

## Example Files

For the Excel file and PDF file as an example, please refer to the attached files.

## Credits

Inspired from Kenji Explains.

