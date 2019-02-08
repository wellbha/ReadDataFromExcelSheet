package readexcelsheet;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.IOException;

public class ExcelUtils {

    static XSSFWorkbook workbook;
    static XSSFSheet sheet;

    @Test
    public static void getRowCount(){
        try {
            workbook = new XSSFWorkbook("src\\test\\Properties\\DataSheet.xlsx");
            sheet = workbook.getSheet("Sheet1");
            int rowCount = sheet.getPhysicalNumberOfRows();
            System.out.println("Number of rows are: " +rowCount);
        } catch (IOException e) {
            System.out.println(e.getMessage());
            System.out.println(e.getCause());
            e.printStackTrace();
        }
    }
    @Test
    public static void getCellData(){
        try {
            workbook = new XSSFWorkbook("src\\test\\Properties\\DataSheet.xlsx");
            sheet = workbook.getSheet("Sheet1");
            String cellData00 = sheet.getRow(0).getCell(0).getStringCellValue();
            String cellData01 = sheet.getRow(0).getCell(1).getStringCellValue();
            String cellData10 = sheet.getRow(1).getCell(0).getStringCellValue();
            double cellData11 = sheet.getRow(1).getCell(1).getNumericCellValue();
            String cellData20 = sheet.getRow(2).getCell(0).getStringCellValue();
            double cellData21 = sheet.getRow(2).getCell(1).getNumericCellValue();
            String cellData30 = sheet.getRow(3).getCell(0).getStringCellValue();
            double cellData31 = sheet.getRow(3).getCell(1).getNumericCellValue();
            System.out.print(cellData00 +"    ");
            System.out.println(cellData01);
            System.out.print(cellData10 +"    ");
            System.out.println(cellData11);
            System.out.print(cellData20 +"    ");
            System.out.println(cellData21);
            System.out.print(cellData30 +"    ");
            System.out.println(cellData31);
        } catch (IOException e) {
            System.out.println(e.getMessage());
            System.out.println(e.getCause());
            e.printStackTrace();
        }
    }
    @Test
    public void getCellDataParameters(int row, int column){
        try {
            workbook = new XSSFWorkbook("src\\test\\Properties\\DataSheet.xlsx");
            sheet = workbook.getSheet("Sheet1");

            String cellData1 = sheet.getRow(row).getCell(column).getStringCellValue();
//            double cellData2 = sheet.getRow(row).getCell(column).getNumericCellValue();

            System.out.print(cellData1 +"    ");
//            System.out.println(cellData2);

        } catch (IOException e) {
            System.out.println(e.getMessage());
            System.out.println(e.getCause());
            e.printStackTrace();
        }
    }
}
