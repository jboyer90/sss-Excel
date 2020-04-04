package com.sss.app;
import com.sss.read.readExcel;
import com.sss.write.writeExcel;

import java.io.IOException;

public class App {
    public static void main(String[] args) throws IOException {
        //writeExcel write = new writeExcel();
    String path = "C:\\Users\\E451Q1\\Desktop\\SSS Holiday Generator\\holiday generated.xlsx";
        //write.writeSingleCellData(path);
       // write.writeMultipleCellData(path);
        //System.out.println("File Created!");
        readExcel read = new readExcel();
        //read.readExcelDataForSingleRow(path);
        read.readExcelDataForEntireSheet(path);
    }
}
