package javaWithExcel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class ReadExcelFile {
    static XSSFWorkbook workbook;
    static XSSFSheet sheet;


    public ReadExcelFile(){
    }
    public ReadExcelFile(String excelPath, String sheetName){
      try {
          workbook = new XSSFWorkbook(excelPath);
          sheet = workbook.getSheet(sheetName);
      }catch (Exception e){
          e.getMessage();
      }
    }
    public  void getExcelRow() {

       int rowCount= sheet.getPhysicalNumberOfRows();
        System.out.println("Number of Rows:" +rowCount);
    }
    public void getCellData(int rowNum, int colNum) {

       String cellData= sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
        System.out.println(cellData);
    }
    //to get a numeric cell value
    public  void getCellDataNumber(int rowNum,int colNum) {

        double cellData= sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
        System.out.println(cellData);
    }
}