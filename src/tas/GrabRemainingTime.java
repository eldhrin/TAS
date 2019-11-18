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
import java.text.ParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;

/**
 *
 * @author fl8328
 */
public class GrabRemainingTime {
     private static Row.MissingCellPolicy xc;
    
    public static void GrabRemainingTime() throws IOException, ParseException{
        
        //connect to the database
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //count is first row where we write the data to
        //find collection TAS
        DBCollection workload = db.getCollection("TAS_WL");
        DBCollection collection = db.getCollection("TAS_PGR");
        DBCursor cursor = workload.find(); 
        //get number of entries in the database (this starts at 1)
        int dbcount = (int)workload.count();
        int dbc = (int)collection.count();
        //base file the program writes to
        FileOutputStream fileOut = new FileOutputStream("megareportKEEP.xlsx");
        //file that the program outputs (creates this file if it does not exist)
        String pathName = "H:\\NetBeansProjects\\TAS\\TAS-VD.xlsx";
        XSSFWorkbook wb = new XSSFWorkbook(pathName);
        int count = 6;
        int i;
        
        DBObject b;
        
        DBCursor cur = collection.find();
        DBObject o = cursor.next();
        
            for(i = 0; i<dbc; i++){
                b = cur.next();
                String lastName = b.get("Supervisor").toString().toUpperCase();
                for(int j = count; j<dbcount+6; j++){  
                    Cell cPGR = wb.getSheetAt(1).getRow(j).getCell(14,xc.CREATE_NULL_AS_BLANK);
                    
                    Cell cLastName = wb.getSheetAt(1).getRow(j).getCell(0,xc.CREATE_NULL_AS_BLANK);
                    String wbLastName = cLastName.getStringCellValue().toUpperCase();
                    
                    if(lastName.equals(wbLastName)){
                        System.out.println(b.get("PGR Allocation").toString());
                        System.out.println(lastName);
                        Double pgr = 0.0;
                        pgr += Double.parseDouble(b.get("PGR Allocation").toString());
                        cPGR.setCellValue(pgr);
                    }
                    
                    if(wbLastName.equals("PETROVSKI")){
                        cPGR.setCellValue(Double.parseDouble(b.get("PGR Allocation").toString())+10);
                    }
                    
                    if(wbLastName.equals("MCDERMOTT")){
                        Cell cFirstName = wb.getSheetAt(1).getRow(j).getCell(1, xc.CREATE_NULL_AS_BLANK);
                        String firstName = cFirstName.getStringCellValue().toUpperCase();
                        if(firstName.equals("CHRIS")){
                            cPGR.setCellValue(0.0);
                        }
                    }
                    
                }
                System.out.println("-------------------------");
            }
                
            
            
          wb.write(fileOut); 
        
        
          fileOut.close();
          //close workbook
          wb.close();
                    
    }
    
        
}

    

