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
import java.util.Date;
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
     Double total = 0.0;
     
     
     //find one from main DB, get the ID and match to research db
     DBCursor compare = collection.find();
     
     XSSFWorkbook wb = new XSSFWorkbook("H:\\NetBeansProjects\\TAS\\TAS-VD.xlsx");  
     
         
        Double count = 0.0; 
        

        for(int i =6; i <= 33; i++){
            DBObject o = compare.next();
            Double research = Double.parseDouble(o.get("Research").toString());
            
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
            
            Double totRes = rcT + ukG + EU + ukC + ukI + ktp + other + sfc + sfcd + pgr + internal + suppint + sfcr;
            
            Double resadj = 0.0;
            
            if(totRes > research){
                resadj = totRes - research;
            }
            else{
                resadj = research - totRes;
            }
            
            
            
            System.out.println(lastName);
            System.out.println("PGR: " + pgr);
            System.out.println("TOTAL RESEARCH :" + totRes);
        
            
            System.out.println("RESEARCH " + research);
                
            if(resadj > 0.0){
                
            Double div = resadj/100;
                
            Double split80 = div*80; 
            Double split20 = div*20;
                
            System.out.println("80 Split: " + split80);
            System.out.println("20 Split: " + split20);
                
            Double ressupport = 0.0;
            ressupport = suppint + split80;
                
            System.out.println(ressupport);
            SuppInt.setCellValue(ressupport); 
                
            phd = phd + split20;
            cphd.setCellValue(phd);
            
            }
            
            total = tAC + tSupp + rcT + ukG + EU + ukC + ukI + ktp + other + sfc + sfcr +  pgr + internal + SuppInt.getNumericCellValue() + sfcr + scholt + scholr + cphd.getNumericCellValue() + otherserv + othersupp + mgmt;
            System.out.println(total);
            
            if(research == 0.0){
                total = total - resadj;
            }
            if(total < 100.00){
                Double remtime = 100 - total;
                
                Double split451 = remtime*45;
                Double split101 = remtime*10;

                System.out.println(split451);
                System.out.println(split101);

                tSupp = (tSupp + split451)/100;
                tsup.setCellValue(tSupp);
                System.out.println("TEACHING SUPPORT SPLIT: " + tSupp);

                scholt = (scholt + split451)/100;
                cscholt.setCellValue(scholt);
                System.out.println("SCHOLARSHIP SPLIT: " + scholt);

                mgmt = (mgmt + split101)/100;
                cmgmt.setCellValue(mgmt);
                System.out.println("MANAGEMENT SPLIT: " + mgmt);
                
            }
            
            Double finalTot = tac.getNumericCellValue() + tsup.getNumericCellValue() + rc.getNumericCellValue() + UKg.getNumericCellValue() + eu.getNumericCellValue() + UKc.getNumericCellValue() + UKi.getNumericCellValue() + KTP.getNumericCellValue() + OTHER.getNumericCellValue() + SFC.getNumericCellValue() + SFCR.getNumericCellValue() + PGR.getNumericCellValue() + INT.getNumericCellValue() + SuppInt.getNumericCellValue() + SFCR.getNumericCellValue() + cscholt.getNumericCellValue() + cscholr.getNumericCellValue() + cphd.getNumericCellValue() + cother.getNumericCellValue() + cothersup.getNumericCellValue() + cmgmt.getNumericCellValue();
            System.out.println("FINAL TOT: " + finalTot);
            Cell ctot = wb.getSheetAt(1).getRow(i).getCell(24, xc.CREATE_NULL_AS_BLANK);
            ctot.setCellValue(finalTot);
            
            Cell chol = wb.getSheetAt(1).getRow(i).getCell(26, xc.CREATE_NULL_AS_BLANK);
            
            Date date = new Date();
            int check = Null.periodChecker(date);
            
            if(check==1){
                Double hol = Double.parseDouble(o.get("sem1").toString());
                chol.setCellValue(hol);
            }
            else if(check==2){
                Double hol = Double.parseDouble(o.get("sem2").toString());
                chol.setCellValue(hol);
            }
            else{
                Double hol = Double.parseDouble(o.get("sem3").toString());
                chol.setCellValue(hol);
            }
           
            
            
            wb.write(fileOut); 
        
            System.out.println("---------------------------------------");

        }
          fileOut.close();
          //close workbook
          wb.close();

                    
                
                //if research allowance = 0, subtract from time remaining
            
            
        
        

    
    }
}

