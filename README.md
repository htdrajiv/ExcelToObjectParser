# ExcelToObjectParser
A library to parse excel file to java object using Apache poi and Google Gson library. 
It can handle nested reference from one class to another and has been implemented using recursion.

In Excel file, we can represent data in simple type or either one of them:
- reference (Composition with Instance of other Class type. eg. excel-data-format1.jpg)
- listReference (Composition with List of Instance of other Class type. eg. excel-data-format2.jpg)

As of now, the header of each column must match with the field name of Java Class. 

Each tab in the excel file represent a Class in Java and the tab name must match with Class Name.
