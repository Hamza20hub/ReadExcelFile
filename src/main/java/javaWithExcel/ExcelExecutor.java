package javaWithExcel;

import java.io.IOException;

public class ExcelExecutor {
    public static void main(String[] args) {
        ReadExcelFile excel = new ReadExcelFile("/Users/anharuzzaman/Desktop/ReadExcleFile/src/AllexcleFiles/Senerio50.xlsx","sheet1");
        excel.getExcelRow();
        //excel.getCellDataNumber(1,1);
        excel.getCellData(1,1);
    }
}
