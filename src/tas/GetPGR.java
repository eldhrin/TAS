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
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author fl8328
 */
public class GetPGR {
    private static Row.MissingCellPolicy xc;
    
    public static void GetPGR() throws IOException, ParseException{
    //connect to local mongodb
    Mongo mongo = new Mongo("localhost", 27017);
    DB tas = mongo.getDB("TAS");
    //find collection TAS
    DBCollection collection = tas.getCollection("TAS_PGR");
        
    //remove all entries from the database
    DBCursor rem = collection.find();
    while(rem.hasNext()){
        collection.remove(rem.next());
    }
    //user chooses directory containing all users tas excel sheets
        
    //read excel file
    XSSFWorkbook wb = new XSSFWorkbook("S:\\Computing\\TAS\\TAS-PGR Supervision.xlsx");                             
                
    for(int i = 13; i < 47; i++){
        Double pgrAdd = 0.0;
         BasicDBObject document = new BasicDBObject();
         
         Cell cCategory = wb.getSheetAt(0).getRow(i).getCell(2, xc.CREATE_NULL_AS_BLANK);
         String category = cCategory.getStringCellValue();
         System.out.println(category);
         Double pgrAllocation = 0.0;
         
         if(category.equals("DOS")){
             pgrAllocation = 5.0;
         }
         else if(category.equals("SUP2")||category.equals("SUP3") || category.equals("STUDY")){
             pgrAllocation = 2.5;
         }
         
         Cell cSupervisor = wb.getSheetAt(0).getRow(i).getCell(3,xc.CREATE_NULL_AS_BLANK);
         String supervisor = cSupervisor.getStringCellValue();
         System.out.println(supervisor);
         String[] sup = supervisor.split(",");
         String supe = sup[0];
         supe = supe.toUpperCase();
                        
         Cell cForP = wb.getSheetAt(0).getRow(i).getCell(4, xc.CREATE_NULL_AS_BLANK);
         String ForP = cForP.getStringCellValue();
         Double pgrAllocationF = 0.0;
         if(ForP.equals("FULL TIME")){
             pgrAllocationF = pgrAllocation;
         }
         else if(ForP.equals("PART TIME")){
             pgrAllocationF = pgrAllocation*.6;
         }
         System.out.println(pgrAllocationF);
        
         BasicDBObject query = new BasicDBObject("Supervisor", supe);
                DBCursor cursor = collection.find(query);
                
                //if the object is found then add to it
                if(cursor.hasNext()){
                    DBObject o = cursor.next();
                    System.out.println(supervisor);
                    System.out.println("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx");
                    System.out.println(o.get("Supervisor").toString());
                    Double pgrOG = Double.parseDouble(o.get("PGR Allocation").toString());
                    pgrOG += pgrAllocationF;
                    document.put("PGR Allocation", pgrOG);
                    collection.update(o, new BasicDBObject("$set",document));
                    
                }
                else{
                    document.put("Supervisor", supe);
                    System.out.println(supervisor);
                    document.put("PGR Allocation", pgrAllocationF);
         
                    collection.insert(document);
                }
        }
    }
    
}
