# Apache-POI-Reading-and-Writing-Excel-file-in-Java

# 1. Basic definitions for Apache POI library
This section briefly describe about basic classes used during Excel Read and Write.

HSSF is prefixed before the class name to indicate operations related to a Microsoft Excel 2003 file.
XSSF is prefixed before the class name to indicate operations related to a Microsoft Excel 2007 file or later.
XSSFWorkbook and HSSFWorkbook are classes which act as an Excel Workbook
HSSFSheet and XSSFSheet are classes which act as an Excel Worksheet
Row defines an Excel row
Cell defines an Excel cell addressed in reference to a row.

# 2. Download Apache POI
Apache POI library is easily available using Maven Dependencies.

pom.xml
  <dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>3.15</version>
  </dependency>
Copy

# 3. Apache POI library – Writing a Simple Excel
The below code shows how to write a simple Excel file using Apache POI libraries. The code uses a 2 dimensional data array to hold the data. The data is written to a XSSFWorkbook object. XSSFSheet is the work sheet being worked on. The code is as shown below:
