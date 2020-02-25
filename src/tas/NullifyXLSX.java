/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tas;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.Mongo;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 *
 * @author fl8328
 */
public class NullifyXLSX {
    
    
    private static Row.MissingCellPolicy xc;
    
    public static XSSFWorkbook nullify() throws IOException{
        
        
        FileOutputStream fileOut = new FileOutputStream("megareportKEEP.xlsx");
        //file that the program outputs (creates this file if it does not exist)
        String pathName = "H:\\NetBeansProjects\\TAS\\TAS-VD.xlsx";
        XSSFWorkbook wb = new XSSFWorkbook(pathName);
        
        
        for(int i=6; i < 34; i++){
            
            for(int j = 3; j<25; j++){
                Cell cell = wb.getSheetAt(1).getRow(i).getCell(j, xc.CREATE_NULL_AS_BLANK);
                cell.setCellValue(0.0);
            }
        }
        
                  wb.write(fileOut); 
        
        
          fileOut.close();
          //close workbook
          wb.close(); 
          
        return null;
    
    }
    
}
    