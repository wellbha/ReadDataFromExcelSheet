package readexcelsheet;

import org.testng.annotations.Test;

public class GetExcelData {

    ExcelUtils data = new ExcelUtils();
    @Test
    public void getRowColumnData(){
        data.getCellDataParameters(0,0);

    }
}
