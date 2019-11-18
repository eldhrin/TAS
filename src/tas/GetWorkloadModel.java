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
public class GetWorkloadModel {
    
    private static Row.MissingCellPolicy xc;
    
    public static void GetWorkloadModel() throws IOException, ParseException{
        //connect to local mongodb
        Mongo mongo = new Mongo("localhost", 27017);
        DB tas = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = tas.getCollection("TAS_PGR");
        DBCollection workload = tas.getCollection("TAS_WL");

        
        DBCursor rem = workload.find();
        while(rem.hasNext()){
            workload.remove(rem.next());
        }
        //user chooses directory containing all users tas excel sheets

        //read excel file
        XSSFWorkbook wb = new XSSFWorkbook("H:\\NetBeansProjects\\TAS\\TAS-Workload Model_18-19v2.xlsx");                             

        for(int i = 5; i < 33; i++){
            BasicDBObject document = new BasicDBObject();
            
            Cell clastName = wb.getSheetAt(0).getRow(i).getCell(0,xc.CREATE_NULL_AS_BLANK);
            String lastName = clastName.getStringCellValue();
            
            Cell cfirstName = wb.getSheetAt(0).getRow(i).getCell(1, xc.CREATE_NULL_AS_BLANK);
            String firstName = cfirstName.getStringCellValue();
            
            Cell cModuleCo1 = wb.getSheetAt(0).getRow(i).getCell(2, xc.CREATE_NULL_AS_BLANK);
            Double moduleCo1 = cModuleCo1.getNumericCellValue();
            moduleCo1 = moduleCo1*5;
            
            Cell cModuleAss1 = wb.getSheetAt(0).getRow(i).getCell(3, xc.CREATE_NULL_AS_BLANK);
            Double moduleAss1 = cModuleAss1.getNumericCellValue();
            moduleAss1 = moduleAss1*5;
            
            Cell cModuleCo2 = wb.getSheetAt(0).getRow(i).getCell(4, xc.CREATE_NULL_AS_BLANK);
            Double moduleCo2 = cModuleCo2.getNumericCellValue();
            moduleCo2 = moduleCo2*5;
            
            Cell cModuleAss2 = wb.getSheetAt(0).getRow(i).getCell(5, xc.CREATE_NULL_AS_BLANK);
            Double moduleAss2 = cModuleAss2.getNumericCellValue();
            moduleAss2 = moduleAss2*5;
            
            Cell cModuleCo3 = wb.getSheetAt(0).getRow(i).getCell(6, xc.CREATE_NULL_AS_BLANK);
            Double moduleCo3 = cModuleCo3.getNumericCellValue();
            moduleCo3 = moduleCo3*5;
            
            Cell cGASept = wb.getSheetAt(0).getRow(i).getCell(7, xc.CREATE_NULL_AS_BLANK);
            Double GASept = cGASept.getNumericCellValue();
            GASept = GASept*5;
            
            Cell cGADec = wb.getSheetAt(0).getRow(i).getCell(8, xc.CREATE_NULL_AS_BLANK);
            Double GADec = cGADec.getNumericCellValue();
            GADec = GADec*5;
            
            Cell cGAMar = wb.getSheetAt(0).getRow(i).getCell(9, xc.CREATE_NULL_AS_BLANK);
            Double GAMar = cGAMar.getNumericCellValue();
            GAMar = GAMar*5;
            
            Cell cGAJun = wb.getSheetAt(0).getRow(i).getCell(10, xc.CREATE_NULL_AS_BLANK);
            Double GAJun = cGAJun.getNumericCellValue();
            GAJun = GAJun*5;
            
            Cell cResearch = wb.getSheetAt(0).getRow(i).getCell(12, xc.CREATE_NULL_AS_BLANK);
            Double research = 0.0;
            research = cResearch.getNumericCellValue();
            research = research*5;
            
            Cell cOtherServicesRendered = wb.getSheetAt(0).getRow(i).getCell(13, xc.CREATE_NULL_AS_BLANK);
            Double otherServicesRendered = 0.0;
            otherServicesRendered = cOtherServicesRendered.getNumericCellValue();
            otherServicesRendered = otherServicesRendered*5;
            
            Cell cSupportOtherServicesRendered = wb.getSheetAt(0).getRow(i).getCell(14, xc.CREATE_NULL_AS_BLANK);
            Double supportOtherServicesRendered = 0.0;
            supportOtherServicesRendered = cSupportOtherServicesRendered.getNumericCellValue();
            supportOtherServicesRendered = supportOtherServicesRendered*5;
            
            Cell cTeachingSupport = wb.getSheetAt(0).getRow(i).getCell(15, xc.CREATE_NULL_AS_BLANK);
            Double teachingSupport = 0.0;
            teachingSupport = cTeachingSupport.getNumericCellValue();
            teachingSupport = teachingSupport*5;
            
            Cell cPGRSupervision = wb.getSheetAt(0).getRow(i).getCell(16, xc.CREATE_NULL_AS_BLANK);
            Double PGRSupervision = 0.0;
            PGRSupervision = cPGRSupervision.getNumericCellValue();
            PGRSupervision = PGRSupervision*5;
            
            Cell cSupportForResearch = wb.getSheetAt(0).getRow(i).getCell(17, xc.CREATE_NULL_AS_BLANK);
            Double supportForResearch = 0.0;
            supportForResearch = cSupportForResearch.getNumericCellValue();
            supportForResearch = supportForResearch*5;
            
            Cell cFurtherEducation = wb.getSheetAt(0).getRow(i).getCell(18, xc.CREATE_NULL_AS_BLANK);
            Double furtherEducation = 0.0;
            furtherEducation = cFurtherEducation.getNumericCellValue();
            furtherEducation = furtherEducation*5;
            
            Cell cMgmt= wb.getSheetAt(0).getRow(i).getCell(19,xc.CREATE_NULL_AS_BLANK);
            Double mgmt = 0.0;
            mgmt = cMgmt.getNumericCellValue();
            mgmt = mgmt*5;
            
            Cell cRemainingTime =  wb.getSheetAt(0).getRow(i).getCell(22, xc.CREATE_NULL_AS_BLANK);
            Double remainingTime = 0.0;
            remainingTime = cRemainingTime.getNumericCellValue();
            
            document.put("lastName", lastName);
            document.put("firstName",firstName);
            document.put("Module Coordination Semester 1", moduleCo1);
            document.put("Module Assist Semester 1", moduleAss1);
            document.put("Module Coordination Semester 2", moduleCo2);
            document.put("Module Assist Semester 2", moduleAss2);
            document.put("Module Coordination Semester 3", moduleCo3);
            document.put("GA September", GASept);
            document.put("GA December", GADec);
            document.put("GA March", GAMar);
            document.put("GA June", GAJun);
            document.put("Research", research);
            document.put("Other Services Rendered", otherServicesRendered);
            document.put("Support for other services rendered", supportOtherServicesRendered);
            document.put("Teaching Support", teachingSupport);
            document.put("PGR Supervision", PGRSupervision);
            document.put("Support for research", supportForResearch);
            document.put("Further Education", furtherEducation);
            document.put("Mgmt", mgmt);
            document.put("Remaining Time", remainingTime);
            
            workload.insert(document);
        }
    }
}
