package tas;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;

import java.io.FileOutputStream;

/**
 *
 * @author fl8328
 */
public class TAS {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
       
        Workbook workbook = new HSSFWorkbook();
        
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(1);
        Cell cell = row.createCell(4);
        
        cell.setCellValue("YO");
        
        try {
            FileOutputStream output = new FileOutputStream("Test.xls");
            workbook.write(output);
            output.close();
        } catch (Exception e){
            e.printStackTrace();
        }
        
    }
    
}
