package tas;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

import java.util.*;
import java.io.*;
import javax.swing.*;


/**
 *
 * @author fl8328
 */
public class TAS {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
       
        //user chooses file
        JFileChooser fileChooser = new JFileChooser();
        int returnValue = fileChooser.showOpenDialog(null);
        //approve file chosen
        if(returnValue == JFileChooser.APPROVE_OPTION){
            
            //TRY CATCH
            //get selected file
            try {
                XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
                Sheet sheet = wb.getSheetAt(0);
                
                //create new file from template file, modifying file goes here
                FileOutputStream fileOut = new FileOutputStream("NEW_TAS.xlsx");
                wb.write(fileOut);
                fileOut.close();
                
                
                //ITERATE THROUGH FILE
//                for(Iterator<Row> rit = sheet.rowIterator(); rit.hasNext();){
//                    Row row = rit.next();
//                    
//                    for(Iterator<Cell> cit = row.cellIterator(); cit.hasNext();){
//                        Cell cell = cit.next();
//                        System.out.println(cell + "\t");
//                    }
//                    System.out.println();
//                }
            } catch (FileNotFoundException ex) {
               ex.printStackTrace();           
            }
            
        }
    }
}
