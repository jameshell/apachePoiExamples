
# Apache POI Examples
Ways to handle MS Office files with Java, going to focus first on .xls and .xlsx MS Excel files using Apache POI.

[![CircleCI](https://img.shields.io/circleci/project/github/jameshell/apachePoiExamples.svg)](https://circleci.com/gh/jameshell/apachePoiExamples) [![GitHub license](https://img.shields.io/github/license/mashape/apistatus.svg)](https://github.com/jameshell/apachePoiExamples) [![Codacy Badge](https://api.codacy.com/project/badge/Grade/471df0437bf2478fb9e857f1085b7b64)](https://www.codacy.com/app/jameshell/apachePoiExamples?utm_source=github.com&amp;utm_medium=referral&amp;utm_content=jameshell/apachePoiExamples&amp;utm_campaign=Badge_Grade)
### Getting Started
Configure your Java project dependencies in order to use Apache POI. You can either use Maven, Gradle or just download the JARs from here: https://poi.apache.org/download.html.

- Maven POM Configuration for HSSF(for Excel 97-2003):

  ```
  <dependencies>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>Insert here desired Version</version>
        </dependency>
    </dependencies>
  ```
  
- Maven POM Configuration for XSSF(for Excel 2007+):

  ```
  <dependencies>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>Insert here desired Version</version>
        </dependency>
    </dependencies>
  ```

#### Microsoft Excel 97-2003
- Creating a simple file

  ```
  Workbook wb = new HSSFWorkbook(); 
        try(OutputStream fileOut = new FileOutputStream("TheNameOfYourFile.xls")){
            wb.write(fileOut);  
        } catch(Exception e){
            System.out.println(e.getMessage());
          }   
  ```
  
- Creating Sheets inside the file
  ```
  Sheet sheet1 = wb.createSheet("Your First Sheet");  
  Sheet sheet2 = wb.createSheet("Your Second Sheet");  
  ```
  
- Setting up the row and col where to insert data
  ```
  Sheet sheet1 = wb.createSheet("Your First Sheet"); 
  Row row = sheet1.createRow(2);   //The integer represents the row/col position.
  Cell cell = row.createCell(2);
  cell.setCellValue("Value 1");    //Value that is to be written in the cell.
  ```
  
  - Inserting Date sample with a specified format in a cell.
  ```
  //Creation helper will aid us in creating a style for the sheet.
    CreationHelper helper = wb.getCreationHelper();
        
   //A style must be created on the workbook in order to be able to insert a date with the desired format.
     CellStyle style = wb.createCellStyle();
     style.setDataFormat(helper.createDataFormat().getFormat("d/m/yy h:mm"));
            
    //We will now create another cell on the same sheet to insert a date input example.
      Row row1 = sheet1.createRow(3);
      Cell cell1 = row1.createCell(3);
            
     //System's current date is created with the new format within the cell.
       cell1.setCellValue(new Date());
       cell1.setCellStyle(style);
  ```
