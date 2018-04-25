package com.mycompany.assignment1rt;


import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.Writer;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Write {

    public void md()throws IOException {

        InputStream ExcelFileToRead = new FileInputStream("C:\\ass1.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

        XSSFWorkbook excel = new XSSFWorkbook();

        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row;
        XSSFCell cell;

        Iterator rows = sheet.rowIterator();
        Writer writer = null;
        File file = new File("C:/Users/ASUS/243121-STIW3054-A172-A1.wiki/List.md");
        writer = new BufferedWriter(new FileWriter(file));

        boolean a = true;
        DataFormatter data = new DataFormatter();
        while (rows.hasNext()) {
            row = (XSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()) {
                cell = (XSSFCell) cells.next();
                String n = data.formatCellValue(cell);

                writer.write("|");
                if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {

                    System.out.print(n + " ");

                    writer.write(n + " ");
                } else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                    System.out.print(n + " ");
                    writer.write(n + " ");
                } else {
                    System.out.print(" ");
                }
            }

            System.out.println();
            
            writer.write("\n");
            if (a == true) {

                writer.write("--|--|--|--\n");
               
                a = false;
            }

        }
        try {
            if (writer != null) {
                writer.close();

            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (writer != null) {
                    writer.close();
                }
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
    }
}
