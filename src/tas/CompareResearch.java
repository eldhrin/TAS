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
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author fl8328
 */
public class CompareResearch {
    
    private static Row.MissingCellPolicy xc;
    
    
    public static void compareResearch() throws IOException, InvalidFormatException, ParseException{
     Mongo mongo = new Mongo("localhost", 27017);
     DB tas = mongo.getDB("TAS");
     //find collection TAS
     DBCollection collection = tas.getCollection("TAS_WL");
     FileOutputStream fileOut = new FileOutputStream("megareportKEEP.xlsx");
     
     
     int dbcount = (int)collection.count();
     
     
     //find one from main DB, get the ID and match to research db
     DBCursor compare = collection.find();
     
     XSSFWorkbook wb = new XSSFWorkbook("H:\\NetBeansProjects\\TAS\\TAS-VD.xlsx");  
     
         
        Double count = 0.0; 
        

        for(int i =6; i <= 33; i++){
            DBObject o = compare.next();
            
            Cell lname = wb.getSheetAt(1).getRow(i).getCell(0, xc.CREATE_NULL_AS_BLANK);
            String lastName = new String();
            lastName = Null.nullString(lname, lastName);
            
            Cell tac = wb.getSheetAt(1).getRow(i).getCell(3, xc.CREATE_NULL_AS_BLANK);
            Double tAC = 0.0;
            tAC = Null.nullDouble(tac, tAC);
            
            Cell tsup = wb.getSheetAt(1).getRow(i).getCell(4, xc.CREATE_NULL_AS_BLANK);
            Double tSupp = 0.0;
            tSupp = Null.nullDouble(tsup, tSupp);
            
            Cell rc = wb.getSheetAt(1).getRow(i).getCell(5, xc.CREATE_NULL_AS_BLANK);
            Double rcT = 0.0;
            rcT = Null.nullDouble(rc, rcT);
            
            Cell UKg = wb.getSheetAt(1).getRow(i).getCell(6, xc.CREATE_NULL_AS_BLANK);
            Double ukG = 0.0;
            ukG = Null.nullDouble(UKg, ukG);
            
            Cell eu = wb.getSheetAt(1).getRow(i).getCell(7, xc.CREATE_NULL_AS_BLANK);
            Double EU = 0.0;
            EU = Null.nullDouble(eu, EU);
            
            Cell UKc = wb.getSheetAt(1).getRow(i).getCell(8, xc.CREATE_NULL_AS_BLANK);
            Double ukC = 0.0;
            ukC = Null.nullDouble(UKc, ukC);
            
            Cell UKi = wb.getSheetAt(1).getRow(i).getCell(9, xc.CREATE_NULL_AS_BLANK);
            Double ukI = 0.0;
            ukI = Null.nullDouble(UKi, ukI);
            
            Cell KTP = wb.getSheetAt(1).getRow(i).getCell(10, xc.CREATE_NULL_AS_BLANK);
            Double ktp = 0.0;
            ktp = Null.nullDouble(KTP, ktp);
            
            Cell OTHER = wb.getSheetAt(1).getRow(i).getCell(11, xc.CREATE_NULL_AS_BLANK);
            Double other = 0.0;
            other = Null.nullDouble(OTHER, other);
            
            Cell SFC = wb.getSheetAt(1).getRow(i).getCell(12, xc.CREATE_NULL_AS_BLANK);
            Double sfc = 0.0;
            sfc = Null.nullDouble(SFC, sfc);
            
            Cell SFCD = wb.getSheetAt(1).getRow(i).getCell(13, xc.CREATE_NULL_AS_BLANK);
            Double sfcd = 0.0;
            sfcd = Null.nullDouble(SFCD, sfcd);
            
            Cell PGR = wb.getSheetAt(1).getRow(i).getCell(14, xc.CREATE_NULL_AS_BLANK);
            Double pgr = 0.0;
            pgr = Null.nullDouble(PGR, pgr);
            
            Cell INT = wb.getSheetAt(1).getRow(i).getCell(15, xc.CREATE_NULL_AS_BLANK);
            Double internal = 0.0;
            internal = Null.nullDouble(INT, internal);
            
            Cell SuppInt = wb.getSheetAt(1).getRow(i).getCell(16, xc.CREATE_NULL_AS_BLANK);
            Double suppint = 0.0;
            suppint = Null.nullDouble(SuppInt, suppint);
            
            Cell SFCR = wb.getSheetAt(1).getRow(i).getCell(17, xc.CREATE_NULL_AS_BLANK);
            Double sfcr = 0.0;
            sfcr = Null.nullDouble(SFCR, sfcr);
            
            Cell cscholt = wb.getSheetAt(1).getRow(i).getCell(18,xc.CREATE_NULL_AS_BLANK);
            Double scholt = 0.0;
            scholt = Null.nullDouble(cscholt, scholt);
            
            Cell cscholr = wb.getSheetAt(1).getRow(i).getCell(19, xc.CREATE_NULL_AS_BLANK);
            Double scholr = 0.0;
            scholr = Null.nullDouble(cscholr, scholr);
            
            Cell cphd = wb.getSheetAt(1).getRow(i).getCell(20, xc.CREATE_NULL_AS_BLANK);
            Double phd = 0.0;
            phd = Null.nullDouble(cphd, phd);
            
            Cell cother = wb.getSheetAt(1).getRow(i).getCell(21, xc.CREATE_NULL_AS_BLANK);
            Double otherserv = 0.0;
            otherserv = Null.nullDouble(cother, otherserv);
            
            Cell cothersup = wb.getSheetAt(1).getRow(i).getCell(22, xc.CREATE_NULL_AS_BLANK);
            Double othersupp = 0.0;
            othersupp = Null.nullDouble(cothersup, othersupp);
            
            Cell cmgmt = wb.getSheetAt(1).getRow(i).getCell(23, xc.CREATE_NULL_AS_BLANK);
            Double mgmt = 0.0;
            mgmt = Null.nullDouble(cmgmt, mgmt);
            
            Double totRes = rcT + ukG + EU + ukC + ukI + ktp + other + sfc + sfcd + internal + suppint + sfcr;
            
            Double total = tAC + tSupp + rcT + ukG + EU + ukC + ukI + ktp + other + sfc + sfcd + internal + suppint + sfc + scholt + scholr + phd + otherserv + othersupp + mgmt;
            System.out.println("TOTAL: " + total);
            System.out.println(lastName);
            System.out.println("PGR: " + pgr);
            System.out.println("TOTAL RESEARCH :" + totRes);
        
            
            String lName = o.get("lastName").toString();
            
            Double research = Double.parseDouble(o.get("Research").toString());
            Double remTime = Double.parseDouble(o.get("Remaining Time").toString());
            System.out.println("RESEARCH " + research);
            Double adjusted = 0.0;
            
            if(research == 0.0){
                adjusted = remTime - totRes;
                if(adjusted < 0.0){
                    System.out.println("ERROR !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                System.out.println("REMAINING TIME: " + remTime);
                System.out.println("NULL " + adjusted);
                
                Double div = adjusted/100;
                
                Double split45 = div*45;
                
                Double split10 = div*10;
                
                System.out.println("nnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn");
                System.out.println(split45);
                System.out.println(split10);
                

                
                tSupp = Null.nullDouble(tsup, tSupp);
                tSupp = tSupp + split45;
                tsup.setCellValue(tSupp);
                
                scholt = Null.nullDouble(cscholt, scholt);
                scholt = scholt + split45;
                cscholt.setCellValue(scholt);
                
                mgmt = Null.nullDouble(cmgmt, mgmt);
                mgmt = mgmt + split10;
                cmgmt.setCellValue(mgmt);
            }
            
            
            else{
                adjusted = research - totRes;
                if(adjusted < 0.0){
                    System.out.println("ADJUSTED " + research + " - " + totRes + " = " + adjusted);
                    System.out.println("ERROR !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                System.out.println("FULL " + adjusted);
                
                Double div = adjusted/100;
                
                Double split80 = div*80; 
                Double split20 = div*20;
                
                System.out.println(split80);
                System.out.println(split20);
                
                Double ressupport = 0.0;
                ressupport = suppint + split80;
                
                System.out.println(ressupport);
                SuppInt.setCellValue(ressupport); 
                
                phd = phd + split20;
                cphd.setCellValue(phd);
                
                System.out.println("REMAINING TIME: " + remTime);
                Double remtime = remTime/100;
                
                Double split451 = remtime*45;
                Double split101 = remtime*10;
                
                System.out.println(split451);
                System.out.println(split101);
               
                tSupp = tSupp + split451;
                tsup.setCellValue(tSupp);
                
                scholt = scholt + split451;
                cscholt.setCellValue(scholt);
                
                mgmt = mgmt + split101;
                cmgmt.setCellValue(mgmt);
                

            }
            total = tAC + tSupp + rcT + ukG + EU + ukC + ukI + ktp + other + sfc + sfcd + pgr + internal + suppint + sfc + scholt + scholr + phd + otherserv + othersupp + mgmt;
            System.out.println("FINAL TOT: " + total);
            
            Cell ctotal = wb.getSheetAt(1).getRow(i).getCell(24, xc.CREATE_NULL_AS_BLANK);
            ctotal.setCellValue(total);
            
            
            wb.write(fileOut); 
        
            System.out.println("---------------------------------------");

        }
          fileOut.close();
          //close workbook
          wb.close();

                    
                
                //if research allowance = 0, subtract from time remaining
            
            
        
        

    
    }
}

