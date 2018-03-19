
package com.mycompany.assignment1rt;



import java.io.BufferedWriter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.Iterator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class assignment243121 {
    
    
    public static final String SAMPLE_XLSX_FILE_PATH = "C:\\ass1.xlsx";
    Writer w=null;
            

    public static void main(String[] args) throws IOException, InvalidFormatException {
       
       try{
           DataFormatter dataFormatter = new DataFormatter();
       
          FileInputStream excelFile = new FileInputStream(new File(SAMPLE_XLSX_FILE_PATH));
          Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0);
            
        Iterator<Row> rowIterator = sheet.rowIterator();
        File file = new File("C:\\syaba.md");
        Writer w = new BufferedWriter(new FileWriter(file));
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

        
           
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "|");
                w.write(cellValue + "|");
            }
            System.out.println();
        }


        workbook.close();
    }
       catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        
          
        }    
    }
