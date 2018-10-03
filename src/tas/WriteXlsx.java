//Adam Lyons 03/10/2018
//

package tas;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.Mongo;
import com.mongodb.MongoClient;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;
import com.mongodb.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JFileChooser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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
        
        DBCursor cursor = collection.find();  
        int dbcount = (int)collection.count();
        XSSFWorkbook wb = new XSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
        wb = new XSSFWorkbook("H:\\NetBeansProjects\\TAS\\megareport.xlsx");
                     
            
            
            for(int j = 0; j < dbcount; j++){
                DBObject o = cursor.next();
                DBObject t = (DBObject) o.get("Teaching");
                DBObject r = (DBObject) o.get("Research");
                DBObject s = (DBObject) o.get("Scholarship");
                DBObject q = (DBObject) o.get("Other");
                
                //INIT
                    Cell tdate = wb.getSheetAt(0).getRow(3).getCell(2);
                    tdate.setCellValue(o.get("date").toString());
                    
                    Cell tschool = wb.getSheetAt(0).getRow(5).getCell(2);
                    tschool.setCellValue(o.get("school").toString());
                    
                //USER
                    Cell tname = wb.getSheetAt(0).getRow(count).getCell(0);
                    tname.setCellValue(o.get("name").toString());
                    
                    Cell tuid = wb.getSheetAt(0).getRow(count).getCell(1);
                    tuid.setCellValue(o.get("uID").toString());
                    
                    //TEACHING
                    Cell tcore = wb.getSheetAt(0).getRow(count).getCell(2);
                    tcore.setCellValue(Double.parseDouble(t.get("core").toString()));
                    
                    Cell tsupport = wb.getSheetAt(0).getRow(count).getCell(3);
                    tsupport.setCellValue(Double.parseDouble(t.get("support").toString()));
                    
                    //RESEARCH
                    Cell tcouncil = wb.getSheetAt(0).getRow(count).getCell(4);
                    tcouncil.setCellValue(Double.parseDouble(r.get("council").toString()));
                    
                    Cell tuk_govt = wb.getSheetAt(0).getRow(count).getCell(5);
                    tuk_govt.setCellValue(Double.parseDouble(r.get("UK_govt").toString()));
                    
                    Cell teu = wb.getSheetAt(0).getRow(count).getCell(6);
                    teu.setCellValue(Double.parseDouble(r.get("EU").toString()));
                    
                    Cell tuk_charity = wb.getSheetAt(0).getRow(count).getCell(7);
                    tuk_charity.setCellValue(Double.parseDouble(r.get("UK_charity").toString()));
                  
                    Cell tuk_industry = wb.getSheetAt(0).getRow(count).getCell(8);
                    tuk_industry.setCellValue(Double.parseDouble(r.get("UK_industry").toString()));
                    
                    Cell tktp_projects = wb.getSheetAt(0).getRow(count).getCell(9);
                    tktp_projects.setCellValue(Double.parseDouble(r.get("KTP_projects").toString()));
                    
                    Cell tother = wb.getSheetAt(0).getRow(count).getCell(10);
                    tother.setCellValue(Double.parseDouble(r.get("other").toString()));
                    
                    Cell tsfc_innovation = wb.getSheetAt(0).getRow(count).getCell(11);
                    tsfc_innovation.setCellValue(Double.parseDouble(r.get("SFC_innovation").toString()));
                    
                    Cell tsfc_rd = wb.getSheetAt(0).getRow(count).getCell(12);
                    tsfc_rd.setCellValue(Double.parseDouble(r.get("SFC_RD").toString()));
                    
                    Cell tpgr_supervision = wb.getSheetAt(0).getRow(count).getCell(13);
                    tpgr_supervision.setCellValue(Double.parseDouble(r.get("PGR_supervision").toString()));
                    
                    Cell tinternal_research = wb.getSheetAt(0).getRow(count).getCell(14);
                    tinternal_research.setCellValue(Double.parseDouble(r.get("internal_research").toString()));
                    
                    Cell tsupport_intext = wb.getSheetAt(0).getRow(count).getCell(15);
                    tsupport_intext.setCellValue(Double.parseDouble(r.get("support_intext").toString()));
                    
                    Cell tsupport_sfc = wb.getSheetAt(0).getRow(count).getCell(16);
                    tsupport_sfc.setCellValue(Double.parseDouble(r.get("support_SFC").toString()));
                    
                    //SCHOLARSHIP
                    Cell tteaching = wb.getSheetAt(0).getRow(count).getCell(17);
                    tteaching.setCellValue(Double.parseDouble(s.get("teaching").toString()));
                    
                    Cell tresearch = wb.getSheetAt(0).getRow(count).getCell(18);
                    tresearch.setCellValue(Double.parseDouble(s.get("research").toString()));
                    
                    Cell tphd = wb.getSheetAt(0).getRow(count).getCell(19);
                    tphd.setCellValue(Double.parseDouble(s.get("PhD").toString()));
                    
                    //OTHER
                    Cell toother = wb.getSheetAt(0).getRow(count).getCell(20);
                    toother.setCellValue(Double.parseDouble(q.get("Other").toString()));
                    
                    Cell tosupport = wb.getSheetAt(0).getRow(count).getCell(21);
                    tosupport.setCellValue(Double.parseDouble(q.get("Osupport").toString()));

                //MISC
                    Cell tmgmt = wb.getSheetAt(0).getRow(count).getCell(22);
                    tmgmt.setCellValue(Double.parseDouble(o.get("Mgmt").toString()));
                    
                    Cell ttotal = wb.getSheetAt(0).getRow(count).getCell(23);
                    ttotal.setCellValue(Double.parseDouble(o.get("Total").toString()));

                    Cell thols = wb.getSheetAt(0).getRow(count).getCell(25);
                    thols.setCellValue(Double.parseDouble(o.get("Hols").toString()));
                
                    //calc TAS RESEARCH
                    Double council = Double.parseDouble(r.get("council").toString());
                    Double UK_govt = Double.parseDouble(r.get("UK_govt").toString());
                    Double EU = Double.parseDouble(r.get("EU").toString());
                    Double UK_charity = Double.parseDouble(r.get("UK_charity").toString());
                    Double UK_industry = Double.parseDouble(r.get("UK_industry").toString());
                    Double KTP_projects = Double.parseDouble(r.get("KTP_projects").toString());
                    Double R_other = Double.parseDouble(r.get("other").toString());
                    Double SFC_innovation= Double.parseDouble(r.get("SFC_innovation").toString());
                    Double SFC_RD = Double.parseDouble(r.get("SFC_RD").toString());
                    Double internal_research = Double.parseDouble(r.get("internal_research").toString());
                    Double support_intext = Double.parseDouble(r.get("support_intext").toString());
                    Double support_SFC = Double.parseDouble(r.get("support_SFC").toString());

                    Double tasresearch = council + UK_govt + EU + UK_charity + UK_industry + KTP_projects + R_other + SFC_innovation + SFC_RD + internal_research + support_intext + support_SFC/100;
                    Cell tas = wb.getSheetAt(0).getRow(count).getCell(26);
                    tas.setCellValue(tasresearch);

                    count++;
            }
             
            wb.write(fileOut);
            fileOut.close();
            wb.close();
        
        mongo.close();
        return null;
    }
        
}
