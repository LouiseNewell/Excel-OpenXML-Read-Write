# Test Excel
.net core app to test reading and writing Excel files using OpenXml

Uses package DocumentFormat.OpenXml

## Overview
Reading and writing Excel files using OpenXml is fairly complex.

When writing to Excel, if any of the XML is poorly formed it may not be possible to the open the spreadsheet.

This code is to test using OpenXml to both read and write Excel files.
It uses various data types with necesary styling for each.
It also demonstrates styling on the header rows, and basic setting of column widths.

## Data Models
**CellView** is a non application specific view of the data cell.

It is used to convert between the Excel cell and fields in the application specific data model.

**TestView** is an application specific view of the data row.

It has a field for each data type to demonstrate how they should be handled.

## Interfaces and Services
**IExcel and SExcel** are the application specific interface and service.

This will need to be modified per application.
It converts data between the application specific and non application specific data models.
It builds the Stylesheet and Columns required specific to the application.
It uses IExcelOpenXML and SExcelOpenXML which are the non application specific interface and service.

**IExcelOpenXML and SExcelOpenXML** are the non application specific interface and service.

This should be reusable for different applications without modification.
