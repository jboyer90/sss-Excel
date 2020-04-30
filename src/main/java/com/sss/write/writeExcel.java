package com.sss.write;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class writeExcel {

    public void writeSingleCellData(String filePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Employee Info");
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
        int row1 = 6;
        int col = 5;
        int[][] dataArray = getRandomDataArray(row1, col);
        for (int i=0;i<row1;i++){
            Row row =sheet.createRow(i);
            row.createCell(0);
            for(int j=0;j<col;j++){
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

        for (int i =0;i<row;i++){
            for (int j=0;j<col;j++){
                dataArray[i][j] = (int)(Math.random()*1000);
            }
        }
        return dataArray;
    }
}
