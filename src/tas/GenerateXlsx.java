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
import static tas.GetXlsx.nullDouble;
import static tas.GetXlsx.nullString;

/**
 *
 * @author fl8328
 */
public class GenerateXlsx {
    
    public static void generateXlsx() throws IOException, InvalidFormatException{
    Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("TAS");
        int dbcount = (int)collection.count();
        DBCursor cursor = collection.find(); 
        
        for(int i = 0; i < dbcount; i++){
            DBObject o = cursor.next();
            DBObject t = (DBObject) o.get("Teaching");
            DBObject r = (DBObject) o.get("Research");
            DBObject s = (DBObject) o.get("Scholarship");
            DBObject q = (DBObject) o.get("Other");
            String name = o.get("name").toString();
            XSSFWorkbook wb = new XSSFWorkbook();
            FileOutputStream fileOut = new FileOutputStream(name + ".xlsx");
         //TRY CATCH
            //get selected file
                
                //GET VARIABLES FROM THE SPREADSHEET AND CONVERT TO STRING/DOUBLE
                //get name, school, date
                Cell cID = wb.getSheetAt(0).getRow(11).getCell(1);
                cID = (Cell) o.get("uID");
                Cell cDate = wb.getSheetAt(0).getRow(8).getCell(1);
                cDate = (Cell) o.get("date");
                Cell cName = wb.getSheetAt(0).getRow(10).getCell(1);
                cName = (Cell) o.get("name");
                Cell cSchool = wb.getSheetAt(0).getRow(12).getCell(1);
                cSchool = (Cell) o.get("school");
                      
                //TEACHING
                Cell cCore = wb.getSheetAt(0).getRow(16).getCell(2);
                cCore.setCellValue(Double.parseDouble(t.get("core").toString()));
                Double core = 0.0;
                core = Double.parseDouble(t.get("core").toString());
                Cell cSupport = wb.getSheetAt(0).getRow(17).getCell(2);
                cSupport.setCellValue(Double.parseDouble(t.get("support").toString()));
                Double support = 0.0;
                support = Double.parseDouble(t.get("support").toString());
                
                //RESEARCH
                Cell cCouncils = wb.getSheetAt(0).getRow(20).getCell(2);
                cCouncils.setCellValue(Double.parseDouble(r.get("council").toString()));
                Cell cUK_govt = wb.getSheetAt(0).getRow(21).getCell(2);
                cUK_govt.setCellValue(Double.parseDouble(r.get("UK_govt").toString()));
                Cell cEU = wb.getSheetAt(0).getRow(22).getCell(2);
                cEU.setCellValue(Double.parseDouble(r.get("EU").toString()));
                Cell cUK_charity = wb.getSheetAt(0).getRow(23).getCell(2);
                cUK_charity.setCellValue(Double.parseDouble(r.get("UK_charity").toString()));
                Cell cUK_industry = wb.getSheetAt(0).getRow(24).getCell(2);
                cUK_industry.setCellValue(Double.parseDouble(r.get("UK_industry").toString()));
                Cell cKTP_projects = wb.getSheetAt(0).getRow(25).getCell(2);
                cKTP_projects.setCellValue(Double.parseDouble(r.get("KTP_projects").toString()));
                Cell cOther = wb.getSheetAt(0).getRow(26).getCell(2);
                cOther.setCellValue(Double.parseDouble(r.get("other").toString()));
                Cell cSFC_innovation = wb.getSheetAt(0).getRow(27).getCell(2);
                cSFC_innovation.setCellValue(Double.parseDouble(r.get("SFC_innovation").toString()));
                Cell cSFC_RD = wb.getSheetAt(0).getRow(28).getCell(2);
                cSFC_RD.setCellValue(Double.parseDouble(r.get("SFC_RD").toString()));
                Cell cPGR_supervision = wb.getSheetAt(0).getRow(29).getCell(2);
                cPGR_supervision.setCellValue(Double.parseDouble(r.get("PGR_supervision").toString()));
                Cell cInternal_research = wb.getSheetAt(0).getRow(30).getCell(2);
                cInternal_research.setCellValue(Double.parseDouble(r.get("internal_research").toString()));
                Cell cSupport_intext= wb.getSheetAt(0).getRow(31).getCell(2);
                cSupport_intext.setCellValue(Double.parseDouble(r.get("support_intext").toString()));
                Cell cSupport_SFC = wb.getSheetAt(0).getRow(32).getCell(2);
                cSupport_SFC.setCellValue(Double.parseDouble(r.get("support_SFC").toString()));
                
                //SCHOLARSHIP
                Cell cTeaching = wb.getSheetAt(0).getRow(34).getCell(2);
                cTeaching.setCellValue(Double.parseDouble(s.get("teaching").toString()));
                Cell cResearch = wb.getSheetAt(0).getRow(35).getCell(2);
                cResearch.setCellValue(Double.parseDouble(s.get("research").toString()));
                Cell cPhD = wb.getSheetAt(0).getRow(36).getCell(2);
                cPhD.setCellValue(Double.parseDouble(s.get("PhD").toString()));
                
                //OTHER
                Cell coOther = wb.getSheetAt(0).getRow(38).getCell(2);
                coOther.setCellValue(Double.parseDouble(q.get("Other").toString()));
                Cell coSupport = wb.getSheetAt(0).getRow(39).getCell(2);
                coSupport.setCellValue(Double.parseDouble(q.get("Osupport").toString()));
                
                //MANAGEMENT
                Cell cMgmt = wb.getSheetAt(0).getRow(41).getCell(2);
                cMgmt.setCellValue(Double.parseDouble(q.get("Mgmt").toString()));
                
                //TOTAL
                Double ctotal = core + support;
                
                //HOLIDAYS
                Cell cHols = wb.getSheetAt(0).getRow(45).getCell(2);
                cHols.setCellValue(Double.parseDouble(o.get("Hols").toString()));
            }
    }
}
