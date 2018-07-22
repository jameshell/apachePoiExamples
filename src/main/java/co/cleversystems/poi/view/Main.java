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
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author Jaime Alonso - Clever Systems
 */
public class Main {
    public static void main(String[] args) throws FileNotFoundException, IOException {
        /* 
        This will create an excel file on the current project directory. The Excel file will be
        in a 97-2003 format since is uses the HSSFWorkbook.
        */
        Workbook wb = new HSSFWorkbook(); 
        try(OutputStream fileOut = new FileOutputStream("TheNameOfYourFile.xls")){
            wb.write(fileOut);  
        } catch(Exception e){
            System.out.println(e.getMessage());
          }                   
    }
}
