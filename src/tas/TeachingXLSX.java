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
import java.util.Calendar;
import java.util.Date;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author fl8328
 */
public class TeachingXLSX {
    
    
     public static void teachingXLSX() throws IOException, InvalidFormatException, ParseException{
         
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("TAS2");
        DBCollection idCollection = db.getCollection("TAS");
        int dbcount = (int)collection.count();
        int idcount = (int)idCollection.count();
        DBCursor cursor = collection.find(); 
        DBCursor idCursor = idCollection.find();
        XSSFWorkbook wb = new XSSFWorkbook("H:\\NetBeansProjects\\TAS\\tas_blank.xlsx");
        FileOutputStream fileOut = null;
        
        Date sem = new Date();
        Calendar cal = Calendar.getInstance();
        cal.setTime(sem);
        
        int check = Null.periodChecker(sem);
        
         System.out.println(check);
         
     
        for(int i = 0; i < dbcount; i++){
            DBObject o = cursor.next();
            DBObject id = idCursor.next();
            String fname = o.get("fir").toString();
            String lname = o.get("sur").toString();
            String fullName = fname + " " + lname;
            BasicDBObject query = new BasicDBObject("name", fullName);
            DBCursor findID = idCollection.find(query);
            String ID = "";
            if(findID.hasNext()){
                ID = id.get("uID").toString();
            }
            
            fileOut = new FileOutputStream("H:\\NetBeansProjects\\TAS\\test\\" + lname + " " + fname + ".xlsx");
     
            
            Cell cName = wb.getSheetAt(0).getRow(10).getCell(1);
            cName.setCellValue(fname + " " + lname);
            
            Cell cSchool = wb.getSheetAt(0).getRow(12).getCell(1);
            cSchool.setCellValue("CSDM");
            
            Cell cID = wb.getSheetAt(0).getRow(11).getCell(1);
            cID.setCellValue(ID);
            
            if(check == 1){
                Cell cCTeach = wb.getSheetAt(0).getRow(16).getCell(2);
                Double module = (Double)o.get("module co-ord");
                Double moduleS = (Double)o.get("assist");
                Double moduleC = module + moduleS;
                cCTeach.setCellValue(moduleC);
                
                Cell cSTeach = wb.getSheetAt(0).getRow(17).getCell(2);
                Double moduleL = (Double)o.get("module supp");
                cSTeach.setCellValue(moduleL);
            }
            
            else if(check == 2){
                Cell cCTeach = wb.getSheetAt(0).getRow(16).getCell(2);
                Double module = (Double)o.get("module co-ord sem2");
                Double moduleS = (Double)o.get("massist sem2");
                Double moduleC = module + moduleS;
                cCTeach.setCellValue(moduleC);
            
            
                Cell cSTeach = wb.getSheetAt(0).getRow(17).getCell(2);
                Double moduleL = (Double)o.get("module supp sem2");
                cSTeach.setCellValue(moduleS);
                
            }
            else{
                Cell cCTeach = wb.getSheetAt(0).getRow(16).getCell(2);
                Double module = (Double)o.get("module co-ord sem3");
                cCTeach.setCellValue(module);
                
                Cell cSTeach = wb.getSheetAt(0).getRow(17).getCell(2);
                Double moduleS = (Double)o.get("module supp sem3");
                cSTeach.setCellValue(moduleS);
            }
            
            Cell rSupp = wb.getSheetAt(0).getRow(31).getCell(2);
            Double rSuppR = (Double)o.get("research");
            rSupp.setCellValue("rSuppR");
            
            Cell rSuppDev = wb.getSheetAt(0).getRow(32).getCell(2);
            Double rSuppRDev = (Double)o.get("research L");
            rSuppDev.setCellValue(rSuppRDev);
            
            
            Cell cPGRS = wb.getSheetAt(0).getRow(29).getCell(2);
            Double pgr = (Double)o.get("pgr");
            cPGRS.setCellValue(pgr);
            
            Cell cSR = wb.getSheetAt(0).getRow(31).getCell(2);
            Double supportR = (Double)o.get("other supp");
            cSR.setCellValue(supportR);
            
            
            Cell cFurther = wb.getSheetAt(0).getRow(36).getCell(2);
            Double further = (Double)o.get("further");
            cFurther.setCellValue(further);
            
            
            Cell cCommOther = wb.getSheetAt(0).getRow(38).getCell(2);
            Double cOther = (Double)o.get("c other");
            cCommOther.setCellValue(cOther);
            
            Cell cCommOtherS = wb.getSheetAt(0).getRow(39).getCell(2);
            Double cOtherS = (Double)o.get("cother supp");
            cCommOtherS.setCellValue(cOtherS);
            
            
            Cell cMgmt = wb.getSheetAt(0).getRow(41).getCell(2);
            Double mgmt = (Double)o.get("mgmt");
            cMgmt.setCellValue(mgmt);
            wb.write(fileOut);
        }   

     }
     
}
