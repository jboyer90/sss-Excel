package com.sss.read;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class readExcel {

    public void readExcelDataForSingleRow(String path) throws IOException {
        // get workbook object
        File file = new File(path);
        try {
            FileInputStream fis = new FileInputStream(file);

            //get workbook object from above stream
            Workbook workbook = WorkbookFactory.create(fis);

            // get sheet from above
            Sheet sheet = workbook.getSheetAt(0);

            // get row from above
            Row row = sheet.getRow(0);

            // get the cell object
            Cell cell = row.getCell(0);
            System.out.println(cell.getStringCellValue());

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
    public void readExcelDataForEntireSheet(String path) throws EncryptedDocumentException, IOException {
        File file = new File(path);
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(file);
            //get workbook object from above stream
            Workbook workbook = WorkbookFactory.create(fis);

            //iterate through the workbook
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()){
                Sheet sheet = sheetIterator.next();
                // row iterator
                Iterator<Row> rowIterator = sheet.iterator();
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()){
                        case BLANK:
                            System.out.print("" + "\t");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        case ERROR:
                            System.out.print("Error" + "\t");
                            break;
                        case FORMULA:
                            System.out.print(cell.getCellFormula() + "\t");
                            break;
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case _NONE:
                            System.out.print("None" + "\t");
                            break;
                        default:
                            break;
                    }
                }
                System.out.println();
            }

        }  catch (IOException e) {
            e.printStackTrace();
        }
    }
}
