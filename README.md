# Apache POI Examples
Ways to handle MS Office files with Java, going to focus first on .xls and .xlsx MS Excel files using Apache POI.

### Getting Started
Configure your Java project dependencies in order to use Apache POI. You can either use Maven, Gradle or just download the JARs from here: https://poi.apache.org/download.html.

- Maven POM Configuration:
  ```
  <dependencies>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
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
