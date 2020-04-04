package com.sss.write;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class writeExcel {
    /*
        1. single data in excel
        2. multiple data row in excel
     */

    public void writeSingleCellData(String filePath) throws IOException {
        /*
        1. create a workbook
        2. create a sheet in workbook above
        3. Create a row in above sheet
        4. Create a cell in above row
        5. Set data inside cell.
         */
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("First Sheet");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Employee ID");

        // write workbook on output stream
        File file = new File(filePath);
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);

        // close the stream
        fos.close();
        workbook.close();
    }
    public void writeMultipleCellData(String filePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("First Sheet");

        // getRandomDataArray[]
        int[][] dataArray = getRandomDataArray(6,5);
        for (int i=0;i<dataArray.length;i++){
            Row row =sheet.createRow(i);
            row.createCell(0);
            for(int j=0;j<dataArray.length;i++){
                Cell cell = row.createCell(j);
                cell.setCellValue(dataArray[i][j]);
            }
        }

        // write workbook on output stream
        File file = new File(filePath);
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);

        // close the stream
        fos.close();
        workbook.close();
    }
    private int[][] getRandomDataArray(int row, int col) {
        int[][] dataArray = new int[row][col];

        for (int i =0;i<dataArray.length;i++){
            for (int j=0;j<dataArray.length;i++){
                dataArray[i][j] = (int)(Math.random()*1000);
            }
        }
        return dataArray;
    }
}
