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
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author fl8328
 */
public class TeachingXLSX {
    
    
     public static void teachingXLSX() throws IOException, InvalidFormatException{
         
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("TAS2");
        int dbcount = (int)collection.count();
        DBCursor cursor = collection.find(); 
        XSSFWorkbook wb = new XSSFWorkbook("H:\\NetBeansProjects\\TAS\\tas_blank.xlsx");
        FileOutputStream fileOut = null;
         
     
        for(int i = 0; i < dbcount; i++){
            DBObject o = cursor.next();
            String fname = o.get("fir").toString();
            String lname = o.get("sur").toString();
            fileOut = new FileOutputStream("H:\\NetBeansProjects\\TAS\\test\\" + fname + " " + lname + ".xlsx");
     
            
            Cell cName = wb.getSheetAt(0).getRow(10).getCell(1);
            cName.setCellValue(fname + " " + lname);
            
            
        }   
     }
}
