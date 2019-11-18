/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tas;

import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.Mongo;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JOptionPane;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author fl8328
 */
public class PutPGR {
    
    public static void PutPGR() throws IOException{
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("TAS_PGR");
        int dbcount = (int)collection.count();
        DBCursor cursor = collection.find(); 
        XSSFWorkbook wb = new XSSFWorkbook("H:\\NetBeansProjects\\TAS\\Book1.xlsx");
        FileOutputStream fileOut = fileOut = new FileOutputStream("H:\\NetBeansProjects\\TAS\\test\\test.xlsx");
        
        for(int i = 0; i < dbcount; i++){
        
            DBObject o = cursor.next();
            System.out.println("cursor");
            
            Cell cSupName = wb.getSheetAt(0).getRow(i).getCell(0);
                cSupName.setCellValue(o.get("Supervisor").toString());
                System.out.println("sup");
                
            Cell cPGR = wb.getSheetAt(0).getRow(i).getCell(1);
                cPGR.setCellValue(o.get("PGR Allocation").toString());
            }
        
        
                wb.write(fileOut);
                wb.close();
    }
    
}
