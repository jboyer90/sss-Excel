package com.sss.read;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import java.io.*;
import java.util.Iterator;

public class readExcel {

    public void readExcelDataForEntireSheet(String path) throws EncryptedDocumentException {

        try {
            File f = new File(path);
            new FileInputStream(f);
            FileInputStream fis;
            fis = new FileInputStream(f);
            //get workbook object from above stream
            Workbook workbook = WorkbookFactory.create(fis);


            //iterate through the workbook
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()){
                Sheet sheet = sheetIterator.next();

                // For loop for role;

                for (Row myRow : sheet) {
                    for (Cell cell : myRow) {

                        switch (cell.getCellType()){
                            case BLANK:
                                System.out.print("Blank" + "\t");
                                break;
                            case BOOLEAN:
                                System.out.print(cell.getBooleanCellValue() + "\t");
                                break;
                            case ERROR:
                                System.out.print(cell.getErrorCellValue() + "\t");
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
                }
                System.out.println();
            }

        }  catch (IOException e) {
            e.printStackTrace();
        }
    }
}

