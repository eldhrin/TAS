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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author fl8328
 */
public class WriteReportWorkload {
    public static XSSFWorkbook writeReport() throws IOException{
        
        //connect to the database
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //count is first row where we write the data to
        //find collection TAS
        DBCollection workload = db.getCollection("TAS_WL");
        DBCollection collection = db.getCollection("TAS_PGR");
        DBCollection research = db.getCollection("RESEARCH");
        DBCursor cursor = workload.find(); 
        DBCursor cur = collection.find();
        DBCursor res = research.find();
        //get number of entries in the database (this starts at 1)
        int dbcount = (int)workload.count();
        int dbc = (int)collection.count();
        int rescount = (int)research.count();
        //base file the program writes to
        FileOutputStream fileOut = new FileOutputStream("megareportKEEP.xlsx");
        //file that the program outputs (creates this file if it does not exist)
        String pathName = "H:\\NetBeansProjects\\TAS\\TAS-VD.xlsx";
        XSSFWorkbook wb = new XSSFWorkbook(pathName);
        int count = 6;
        double pt5 = 1/.5;
        double pt8 = 1/.8;

            for(int j = count; j < 34; j++){
                       
                DBObject o = cursor.next();
//                DBObject r = res.next();
                

                Cell cLastName = wb.getSheetAt(1).getRow(j).getCell(0);
                String lastN = o.get("lastName").toString();
                cLastName.setCellValue(lastN);
                System.out.println(lastN);

                Cell cFirstName = wb.getSheetAt(1).getRow(j).getCell(1);
                cFirstName.setCellValue(o.get("firstName").toString());

                Cell cTeachingActivity = wb.getSheetAt(1).getRow(j).getCell(3);
                Double moduleCo1 = Double.parseDouble(o.get("Module Coordination Semester 1").toString());
                Double moduleAss1 = Double.parseDouble(o.get("Module Assist Semester 1").toString());
                Double moduleCo2 = Double.parseDouble(o.get("Module Coordination Semester 2").toString());
                Double moduleAss2 = Double.parseDouble(o.get("Module Assist Semester 2").toString());
                Double moduleCo3 = Double.parseDouble(o.get("Module Coordination Semester 3").toString());
                Double GASept = Double.parseDouble(o.get("GA September").toString());
                Double GADec = Double.parseDouble(o.get("GA December").toString());
                Double GAMar = Double.parseDouble(o.get("GA March").toString());
                Double GAJun = Double.parseDouble(o.get("GA June").toString());

                Double moduleCo19 = moduleCo1*.9;
                Double moduleCo11 = moduleCo1*.1;

                Double moduleCo29 = moduleCo2*.9;
                Double moduleCo21 = moduleCo2*.1;

                Double moduleCo39 = moduleCo3*.9;
                Double moduleCo31 = moduleCo3*.1;

                Double GASep9 = GASept*.9;
                Double GASept1 = GASept*.1;

                Double GADec9 = GADec*.9;
                Double GADec1 = GADec*.1;

                Double GAMar9 = GAMar*.9;
                Double GAMar1 = GAMar*.1;

                Double GAJun9 = GAJun*.9;
                Double GAJun1 = GAJun*.1;
                
                if(o.get("lastName").toString().equals("McDonald") || o.get("lastName").toString().equals("McGlone")){
                    Double teachingCore = moduleCo19+moduleCo29+moduleCo39+GASep9+GADec9+GAMar9+GAJun9+moduleAss1+moduleAss2;
                    cTeachingActivity.setCellValue(teachingCore*pt5);

                    Cell cTeachingSupport = wb.getSheetAt(1).getRow(j).getCell(4);
                    Double teachingSupport = moduleCo11+moduleCo21+moduleCo31+GASept1+GADec1+GAMar1+GAJun1;
                    cTeachingSupport.setCellValue(teachingSupport*pt5);               

                    Cell cFurtherEducation = wb.getSheetAt(1).getRow(j).getCell(20);
                    Double furtherEducation = Double.parseDouble(o.get("Further Education").toString());
                    cFurtherEducation.setCellValue(furtherEducation*pt5);

                    Cell cOtherServicesRendered = wb.getSheetAt(1).getRow(j).getCell(21);
                    Double otherServicesRendered = Double.parseDouble(o.get("Other Services Rendered").toString());
                    cOtherServicesRendered.setCellValue(otherServicesRendered*pt5);

                    Cell cSupportforOtherServices = wb.getSheetAt(1).getRow(j).getCell(22);
                    Double supportForOtherServices = Double.parseDouble(o.get("Support for other services rendered").toString());
                    cSupportforOtherServices.setCellValue(supportForOtherServices*pt5);

                    Cell cMgmt = wb.getSheetAt(1).getRow(j).getCell(23);
                    Double mgmt = Double.parseDouble(o.get("Mgmt").toString());
                    cMgmt.setCellValue(mgmt*pt5);
                }
                
                else if (o.get("lastName".toString()).equals("Lothian")){    
                    Double teachingCore = moduleCo19+moduleCo29+moduleCo39+GASep9+GADec9+GAMar9+GAJun9+moduleAss1+moduleAss2;
                    cTeachingActivity.setCellValue(teachingCore*pt8);

                    Cell cTeachingSupport = wb.getSheetAt(1).getRow(j).getCell(4);
                    Double teachingSupport = moduleCo11+moduleCo21+moduleCo31+GASept1+GADec1+GAMar1+GAJun1;
                    cTeachingSupport.setCellValue(teachingSupport*pt8);   

                    Cell cFurtherEducation = wb.getSheetAt(1).getRow(j).getCell(20);
                    Double furtherEducation = Double.parseDouble(o.get("Further Education").toString());
                    cFurtherEducation.setCellValue(furtherEducation*pt8);

                    Cell cOtherServicesRendered = wb.getSheetAt(1).getRow(j).getCell(21);
                    Double otherServicesRendered = Double.parseDouble(o.get("Other Services Rendered").toString());
                    cOtherServicesRendered.setCellValue(otherServicesRendered*pt8);

                    Cell cSupportforOtherServices = wb.getSheetAt(1).getRow(j).getCell(22);
                    Double supportForOtherServices = Double.parseDouble(o.get("Support for other services rendered").toString());
                    cSupportforOtherServices.setCellValue(supportForOtherServices*pt8);

                    Cell cMgmt = wb.getSheetAt(1).getRow(j).getCell(23);
                    Double mgmt = Double.parseDouble(o.get("Mgmt").toString());
                    cMgmt.setCellValue(mgmt*pt8);
                }
                
                else{
                
                    Double teachingCore = moduleCo19+moduleCo29+moduleCo39+GASep9+GADec9+GAMar9+GAJun9+moduleAss1+moduleAss2;
                    cTeachingActivity.setCellValue(teachingCore);

                    Cell cTeachingSupport = wb.getSheetAt(1).getRow(j).getCell(4);
                    Double teachingSupport = moduleCo11+moduleCo21+moduleCo31+GASept1+GADec1+GAMar1+GAJun1;
                    cTeachingSupport.setCellValue(teachingSupport);
                                         
                    Cell cFurtherEducation = wb.getSheetAt(1).getRow(j).getCell(20);
                    Double furtherEducation = Double.parseDouble(o.get("Further Education").toString());
                    cFurtherEducation.setCellValue(furtherEducation);

                    Cell cOtherServicesRendered = wb.getSheetAt(1).getRow(j).getCell(21);
                    Double otherServicesRendered = Double.parseDouble(o.get("Other Services Rendered").toString());
                    cOtherServicesRendered.setCellValue(otherServicesRendered);

                    Cell cSupportforOtherServices = wb.getSheetAt(1).getRow(j).getCell(22);
                    Double supportForOtherServices = Double.parseDouble(o.get("Support for other services rendered").toString());
                    cSupportforOtherServices.setCellValue(supportForOtherServices);

                    Cell cMgmt = wb.getSheetAt(1).getRow(j).getCell(23);
                    Double mgmt = Double.parseDouble(o.get("Mgmt").toString());
                    cMgmt.setCellValue(mgmt);
                
                }
                
            }

            wb.write(fileOut); 
        
        
          fileOut.close();
          //close workbook
          wb.close();
        return null;
    }
                     
}
