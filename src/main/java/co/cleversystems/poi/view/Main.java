/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package co.cleversystems.poi.view;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Jaime Alonso - Clever Systems ©
 */
public class Main {
    public static void main(String[] args) throws FileNotFoundException, IOException {
        /* 
        This will create an excel file on the current project directory. The Excel file will be
        in a 97-2003 (.xls) format since it uses the HSSFWorkbook().
        */
       //Workbook wb = new HSSFWorkbook(); 
        
        
        /*
        This will create an excel file on the current project directory. The Excel file will be 
        in a 2007+ (.xlsx)format since is uses XSSFWorkbook().
        */
        Workbook wb= new XSSFWorkbook();
        
        try(OutputStream fileOut = new FileOutputStream("TheNameOfYourFile.xlsx")){
            
            //The object Sheet will create allow the creation of sheets.
            Sheet sheet1 = wb.createSheet("Your First Sheet");
            Sheet sheet2 = wb.createSheet("Your Second Sheet"); 
            
            /*On the previously created sheet set the row and cell where you want to write data.
              Note that both create row and cell method receive an integer.
            */
            Row row = sheet1.createRow(2);
            Cell cell = row.createCell(2);
            
            //Finally set the value that will be created in the cell.
            cell.setCellValue("Value 1"); 
            
            
            wb.write(fileOut);  
        } catch(Exception e){
            System.out.println(e.getMessage());
          }                   
    }
}
