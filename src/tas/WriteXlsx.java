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
        //find collection TAS
        DBCollection collection = db.getCollection("TAS");
        DBObject dbo = (DBObject)collection.findOne();
        DBCursor cursor = collection.find();
        
        DBObject teaching = (DBObject)dbo.get("Teaching");
        DBObject research = (DBObject)dbo.get("Research");
        DBObject other = (DBObject)dbo.get("Other");
        DBObject scholarship = (DBObject)dbo.get("Scholarship");
        
        XSSFWorkbook wb = new XSSFWorkbook("H:\\NetBeansProjects\\TAS\\test\\megareport.xlsx");
        
         Cell date = wb.getSheetAt(0).getRow(3).getCell(2);
         date.setCellValue(dbo.get("date").toString());
         
         Cell school = wb.getSheetAt(0).getRow(5).getCell(2);
         school.setCellValue(dbo.get("school").toString());
         
         Cell name = wb.getSheetAt(0).getRow(9).getCell(0);
         name.setCellValue(dbo.get("name").toString());
         
         Cell tcore = wb.getSheetAt(0).getRow(9).getCell(1);
         tcore.setCellValue(teaching.get("core").toString());
         
         Cell tsupport = wb.getSheetAt(0).getRow(9).getCell(2);
         tsupport.setCellValue(teaching.get("support").toString());
         
         Cell rcouncils = wb.getSheetAt(0).getRow(9).getCell(3);
         rcouncils.setCellValue(research.get("council").toString());
         
         Cell ruk_govt = wb.getSheetAt(0).getRow(9).getCell(4);
         ruk_govt.setCellValue(research.get("UK_govt").toString());
         
         Cell reu = wb.getSheetAt(0).getRow(9).getCell(5);
         reu.setCellValue(research.get("EU").toString());
         
         Cell ruk_charity = wb.getSheetAt(0).getRow(9).getCell(6);
         ruk_charity.setCellValue(research.get("UK_charity").toString());
         
         Cell ruk_industry = wb.getSheetAt(0).getRow(9).getCell(7);
         ruk_industry.setCellValue(research.get("UK_industry").toString());
         
         Cell rktp_projects = wb.getSheetAt(0).getRow(9).getCell(8);
         rktp_projects.setCellValue(research.get("KTP_projects").toString());
         
         Cell rother = wb.getSheetAt(0).getRow(9).getCell(9);
         rother.setCellValue(research.get("other").toString());
         
         Cell rsfc_innovation = wb.getSheetAt(0).getRow(9).getCell(10);
         rsfc_innovation.setCellValue(research.get("SFC_innovation").toString());
         
         Cell rsfc_rd = wb.getSheetAt(0).getRow(9).getCell(11);
         rsfc_rd.setCellValue(research.get("SFC_RD").toString());
         
         Cell rpgr_supervision = wb.getSheetAt(0).getRow(9).getCell(12);
         rpgr_supervision.setCellValue(research.get("PGR_supervision").toString());
         
         Cell rinternal_research = wb.getSheetAt(0).getRow(9).getCell(13);
         rinternal_research.setCellValue(research.get("internal_research").toString());
         
         Cell rsupport_intext = wb.getSheetAt(0).getRow(9).getCell(14);
         rsupport_intext.setCellValue(research.get("support_intext").toString());
         
         Cell rsupport_sfc = wb.getSheetAt(0).getRow(9).getCell(15);
         rsupport_sfc.setCellValue(research.get("support_SFC").toString());
         
         Cell rteaching = wb.getSheetAt(0).getRow(9).getCell(16);
         rteaching.setCellValue(scholarship.get("teaching").toString());
         
         Cell rresearch = wb.getSheetAt(0).getRow(9).getCell(17);
         rresearch.setCellValue(scholarship.get("research").toString());
         
         Cell rphd = wb.getSheetAt(0).getRow(9).getCell(18);
         rphd.setCellValue(scholarship.get("PhD").toString());
         
         Cell oother = wb.getSheetAt(0).getRow(9).getCell(19);
         oother.setCellValue(other.get("Other").toString());
         
         Cell osupport = wb.getSheetAt(0).getRow(9).getCell(20);
         osupport.setCellValue(other.get("Osupport").toString());
         
         Cell mgmt = wb.getSheetAt(0).getRow(9).getCell(21);
         mgmt.setCellValue(dbo.get("Mgmt").toString());
         
         Cell total = wb.getSheetAt(0).getRow(9).getCell(22);
         total.setCellValue(dbo.get("Total").toString());
         
         Cell hols = wb.getSheetAt(0).getRow(9).getCell(25);
         hols.setCellValue(dbo.get("Hols").toString());
         
        int count = 0;
        
//        while(cursor.hasNext()){
//            
//        
//        
//        //String name = dbo.get("name").toString();
//        
//                    
//        count++;
//        //while has next, add to megareport
//        }
//       
        FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
        wb.write(fileOut);
        fileOut.close();
        mongo.close();
        return null;
    }
}
