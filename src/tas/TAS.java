package tas;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.Mongo;
import com.mongodb.MongoClient;
import com.mongodb.MongoClientURI;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;

import org.json.*;

import java.util.*;
import java.io.*;
import javax.swing.*;
import org.apache.commons.codec.binary.StringUtils;
import org.bson.Document;


/**
 *
 * @author fl8328
 */
public class TAS {
    
    public static Double nullDouble(Cell c, Double d){
        if(c == null){
            d = 0.0;
        }
        else{
            d = Double.parseDouble(c.toString());
        }
        return d;
    }
    
    public static String nullString(Cell c, String s){
        if(c == null){
            s = "";
        }
        else{
            s = c.toString();
        }
        return s;
    }
    

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, JSONException {
       
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        DBCollection collection = db.getCollection("TAS");
//        DBCursor cursor = collection.find();
//        while(cursor.hasNext()){
//            System.out.println(cursor.next());
//        }
        //user chooses file
        JFileChooser fileChooser = new JFileChooser();
        int returnValue = fileChooser.showOpenDialog(null);
        //approve file chosen
        if(returnValue == JFileChooser.APPROVE_OPTION){
            
            
            //TRY CATCH
            //get selected file
            try {
                
                
                XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
                
                Cell cID = wb.getSheetAt(0).getRow(11).getCell(1);
                String cid = new String();
                        cid = nullString(cID,cid);
                Cell cDate = wb.getSheetAt(0).getRow(8).getCell(1);
                String cdate = new String();
                        cdate = nullString(cDate, cdate);
                Cell cName = wb.getSheetAt(0).getRow(10).getCell(1);
                String cname = new String();
                        cname = nullString(cName, cname);
                Cell cSchool = wb.getSheetAt(0).getRow(12).getCell(1);
                String cschool = new String();
                        cschool = nullString(cSchool, cschool);
                
                Cell cCore = wb.getSheetAt(0).getRow(16).getCell(2);
                Double ccore = 0.0;
                        ccore = nullDouble(cCore, ccore);
                Cell cSupport = wb.getSheetAt(0).getRow(17).getCell(2);
                Double csupport = 0.0;
                        csupport = nullDouble(cSupport, csupport);
                
                Cell cCouncils = wb.getSheetAt(0).getRow(20).getCell(2);
                Double ccouncils = 0.0;
                        ccouncils = nullDouble(cCouncils, ccouncils);
                Cell cUK_govt = wb.getSheetAt(0).getRow(21).getCell(2);
                Double cuk_govt = 0.0;
                        cuk_govt = nullDouble(cUK_govt, cuk_govt);
                Cell cEU = wb.getSheetAt(0).getRow(22).getCell(2);
                Double ceu = 0.0;
                        ceu = nullDouble(cEU, ceu);
                Cell cUK_charity = wb.getSheetAt(0).getRow(23).getCell(2);
                Double cuk_charity = 0.0;
                        cuk_charity = nullDouble(cUK_charity, cuk_charity);
                Cell cUK_industry = wb.getSheetAt(0).getRow(24).getCell(2);
                Double cuk_industry = 0.0; 
                        cuk_industry = nullDouble(cUK_industry, cuk_industry);
                Cell cKTP_projects = wb.getSheetAt(0).getRow(25).getCell(2);
                Double cktp_projects = 0.0;
                        cktp_projects = nullDouble(cKTP_projects, cktp_projects);
                Cell cOther = wb.getSheetAt(0).getRow(26).getCell(2);
                Double cother = 0.0;
                        cother = nullDouble(cOther, cother);
                Cell cSFC_innovation = wb.getSheetAt(0).getRow(27).getCell(2);
                Double csfc_innovation = 0.0;
                        csfc_innovation = nullDouble(cSFC_innovation, csfc_innovation);
                Cell cSFC_RD = wb.getSheetAt(0).getRow(28).getCell(2);
                Double csfc_rd = 0.0;
                        csfc_rd = nullDouble(cSFC_RD, csfc_rd);
                Cell cPGR_supervision = wb.getSheetAt(0).getRow(29).getCell(2);
                Double cpgr_supervision = 0.0;
                        cpgr_supervision = nullDouble(cPGR_supervision, cpgr_supervision);
                Cell cInternal_research = wb.getSheetAt(0).getRow(30).getCell(2);
                Double cinternal_research = 0.0;
                        cinternal_research = nullDouble(cInternal_research, cinternal_research);
                Cell cSupport_intext= wb.getSheetAt(0).getRow(31).getCell(2);
                Double csupport_intext = 0.0;
                        csupport = nullDouble(cSupport_intext, csupport);
                Cell cSupport_SFC = wb.getSheetAt(0).getRow(32).getCell(2);
                Double csupport_sfc = 0.0;
                        csupport = nullDouble(cSupport_SFC, csupport);
                
                Cell cTeaching = wb.getSheetAt(0).getRow(34).getCell(2);
                Double cteaching = 0.0;
                    cteaching = nullDouble(cTeaching, cteaching);
                Cell cResearch = wb.getSheetAt(0).getRow(35).getCell(2);
                Double cresearch = 0.0;
                        cresearch = nullDouble(cResearch, cresearch);
                Cell cPhD = wb.getSheetAt(0).getRow(36).getCell(2);
                Double cphd = 0.0;
                        cphd = nullDouble(cPhD, cphd);
                
                Cell coOther = wb.getSheetAt(0).getRow(38).getCell(2);
                Double coother = 0.0;
                        coother = nullDouble(coOther, coother);
                Cell coSupport = wb.getSheetAt(0).getRow(39).getCell(2);
                Double cosupport = 0.0;
                        cosupport = nullDouble(coSupport, cosupport);
                
                Cell cMgmt = wb.getSheetAt(0).getRow(41).getCell(2);
                Double cmgmt = 0.0;
                        cmgmt = nullDouble(cMgmt, cmgmt);
                
                Cell cTotal = wb.getSheetAt(0).getRow(43).getCell(2);
                Double ctotal = 0.0;
                        ctotal = nullDouble(cTotal, ctotal);
                
                Cell cHols = wb.getSheetAt(0).getRow(45).getCell(2);
                Double chols = 0.0;
                        chols = nullDouble(cHols, chols);
                
                BasicDBObject document = new BasicDBObject();
                document.put("_id", cid);
                document.put("date", cdate);
                document.put("name", cname);
                document.put("school", cschool);
                
                BasicDBObject documentTeach = new BasicDBObject();
                documentTeach.put("core", ccore);
                documentTeach.put("support", csupport);
                document.put("Teaching", documentTeach);

                BasicDBObject documentResearch = new BasicDBObject();
                documentResearch.put("UK_govt", cuk_govt);
                documentResearch.put("EU", ceu);
                documentResearch.put("UK_charity", cuk_charity);
                documentResearch.put("UK_industry", cuk_industry);
                documentResearch.put("KTP_projects", cktp_projects);
                documentResearch.put("other", cother);
                documentResearch.put("SFC_innovation", csfc_innovation);
                documentResearch.put("SFC_RD", csfc_rd);
                documentResearch.put("PGR_supervision", cpgr_supervision);
                documentResearch.put("internal_research", cinternal_research);
                documentResearch.put("support_intext", csupport_intext);
                documentResearch.put("support_SFC", csupport_sfc);
                document.put("Research", documentResearch);
                
                BasicDBObject documentSchol = new BasicDBObject();
                documentSchol.put("teaching", cteaching);
                documentSchol.put("research", cresearch);
                documentSchol.put("PhD", cphd);
                document.put("Scholarship", documentSchol);
                
                BasicDBObject documentOther = new BasicDBObject();
                documentOther.put("Other", coother);
                documentOther.put("Osupport", cosupport);
                document.put("Other", documentOther);
                
                document.put("Mgmt", cmgmt);
                document.put("Total", ctotal);
                document.put("Hols", chols);
                
                collection.insert(document);
                System.out.println(document);
                

            } catch (FileNotFoundException ex) {
               ex.printStackTrace();           
            }
        }
        
        
    }
    
}
