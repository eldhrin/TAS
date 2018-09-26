package tas;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.Mongo;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JFileChooser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;
/**
 *
 * @author fl8328
 */
public class WriteXlsx {
    
    
    public static XSSFWorkbook writeXlsx() throws IOException{
        
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        int count = 9;
        //find collection TAS
        DBCollection collection = db.getCollection("TAS");
        DBObject dbo = (DBObject)collection.findOne();
        DBCursor cursor = collection.find();
        int dbcount = (int)collection.count();
        XSSFWorkbook wb = new XSSFWorkbook();
        
        
        if(cursor.hasNext()){
        for(int i = 0; i < dbcount; i ++){
        DBObject teaching = (DBObject)dbo.get("Teaching");
        DBObject research = (DBObject)dbo.get("Research");
        DBObject other = (DBObject)dbo.get("Other");
        DBObject scholarship = (DBObject)dbo.get("Scholarship");

            wb = new XSSFWorkbook("H:\\NetBeansProjects\\TAS\\test\\megareport.xlsx");

             Cell date = wb.getSheetAt(0).getRow(3).getCell(2);
             date.setCellValue(dbo.get("date").toString());

             Cell school = wb.getSheetAt(0).getRow(5).getCell(2);
             school.setCellValue(dbo.get("school").toString());

             Cell name = wb.getSheetAt(0).getRow(count).getCell(0);
             name.setCellValue(dbo.get("name").toString());

             Cell tuid = wb.getSheetAt(0).getRow(count).getCell(1);
             tuid.setCellValue(dbo.get("uID").toString());

             Cell tcore = wb.getSheetAt(0).getRow(count).getCell(2);
             tcore.setCellValue(teaching.get("core").toString());

             Cell tsupport = wb.getSheetAt(0).getRow(count).getCell(3);
             tsupport.setCellValue(teaching.get("support").toString());

             Cell rcouncils = wb.getSheetAt(0).getRow(count).getCell(4);
             rcouncils.setCellValue(research.get("council").toString());

             Cell ruk_govt = wb.getSheetAt(0).getRow(count).getCell(5);
             ruk_govt.setCellValue(research.get("UK_govt").toString());

             Cell reu = wb.getSheetAt(0).getRow(count).getCell(6);
             reu.setCellValue(research.get("EU").toString());

             Cell ruk_charity = wb.getSheetAt(0).getRow(count).getCell(7);
             ruk_charity.setCellValue(research.get("UK_charity").toString());

             Cell ruk_industry = wb.getSheetAt(0).getRow(count).getCell(8);
             ruk_industry.setCellValue(research.get("UK_industry").toString());

             Cell rktp_projects = wb.getSheetAt(0).getRow(count).getCell(9);
             rktp_projects.setCellValue(research.get("KTP_projects").toString());

             Cell rother = wb.getSheetAt(0).getRow(count).getCell(10);
             rother.setCellValue(research.get("other").toString());

             Cell rsfc_innovation = wb.getSheetAt(0).getRow(count).getCell(11);
             rsfc_innovation.setCellValue(research.get("SFC_innovation").toString());

             Cell rsfc_rd = wb.getSheetAt(0).getRow(count).getCell(12);
             rsfc_rd.setCellValue(research.get("SFC_RD").toString());

             Cell rpgr_supervision = wb.getSheetAt(0).getRow(count).getCell(13);
             rpgr_supervision.setCellValue(research.get("PGR_supervision").toString());

             Cell rinternal_research = wb.getSheetAt(0).getRow(count).getCell(14);
             rinternal_research.setCellValue(research.get("internal_research").toString());

             Cell rsupport_intext = wb.getSheetAt(0).getRow(count).getCell(15);
             rsupport_intext.setCellValue(research.get("support_intext").toString());

             Cell rsupport_sfc = wb.getSheetAt(0).getRow(count).getCell(16);
             rsupport_sfc.setCellValue(research.get("support_SFC").toString());

             Cell rteaching = wb.getSheetAt(0).getRow(count).getCell(17);
             rteaching.setCellValue(scholarship.get("teaching").toString());

             Cell rresearch = wb.getSheetAt(0).getRow(count).getCell(18);
             rresearch.setCellValue(scholarship.get("research").toString());

             Cell rphd = wb.getSheetAt(0).getRow(count).getCell(19);
             rphd.setCellValue(scholarship.get("PhD").toString());

             Cell oother = wb.getSheetAt(0).getRow(count).getCell(20);
             oother.setCellValue(other.get("Other").toString());

             Cell osupport = wb.getSheetAt(0).getRow(count).getCell(21);
             osupport.setCellValue(other.get("Osupport").toString());

             Cell mgmt = wb.getSheetAt(0).getRow(count).getCell(22);
             mgmt.setCellValue(Double.parseDouble(dbo.get("Mgmt").toString()));

             Cell total = wb.getSheetAt(0).getRow(count).getCell(23);
             total.setCellValue(dbo.get("Total").toString());

             Cell hols = wb.getSheetAt(0).getRow(count).getCell(25);
             hols.setCellValue(dbo.get("Hols").toString());           
             
             //calc
             Double council = Double.parseDouble(research.get("council").toString());
             Double UK_govt = Double.parseDouble(research.get("UK_govt").toString());
             Double EU = Double.parseDouble(research.get("EU").toString());
             Double UK_charity = Double.parseDouble(research.get("UK_charity").toString());
             Double UK_industry = Double.parseDouble(research.get("UK_industry").toString());
             Double KTP_projects = Double.parseDouble(research.get("KTP_projects").toString());
             Double R_other = Double.parseDouble(research.get("other").toString());
             Double SFC_innovation= Double.parseDouble(research.get("SFC_innovation").toString());
             Double SFC_RD = Double.parseDouble(research.get("SFC_RD").toString());
             Double internal_research = Double.parseDouble(research.get("internal_research").toString());
             Double support_intext = Double.parseDouble(research.get("support_intext").toString());
             Double support_SFC = Double.parseDouble(research.get("support_SFC").toString());
             
             Double tasresearch = council + UK_govt + EU + UK_charity + UK_industry + KTP_projects + R_other + SFC_innovation + SFC_RD + internal_research + support_intext + support_SFC;
             Cell tas = wb.getSheetAt(0).getRow(count).getCell(26);
             tas.setCellValue(tasresearch);

        }
        }
        FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
        wb.write(fileOut);
        fileOut.close();
        System.out.println(count);
             
        count++;
        System.out.println(count);
        mongo.close();
        return null;
        }
    }

