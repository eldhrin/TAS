package tas;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.Mongo;
import com.mongodb.MongoClient;
import com.mongodb.MongoClientURI;
import org.apache.poi.xssf.usermodel.*;

import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import com.mongodb.client.MongoIterable;
import com.mongodb.util.JSON;

import org.json.*;

import java.util.*;
import java.io.*;
import java.net.UnknownHostException;
import javax.swing.*;
import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.bson.Document;
import org.bson.types.ObjectId;


/**
 *
 * @author fl8328
 */
public class TAS {
    
    //if DoubleCell == blank, d = 0.0
    public static Double nullDouble(Cell c, Double d){
        if(c == null){
            d = 0.0;
        }
        //if DoubleCell != blank, d = value of cell
        else{
            d = Double.parseDouble(c.toString());
        }
        return d;
    }
    
    //if StringCell == blank, s = ""
    public static String nullString(Cell c, String s){
        if(c == null){
            s = "";
        }
        //if StringCell != blank, s = value of cell
        else{
            s = c.toString();
        }
        return s;
    }
    
    private static MissingCellPolicy xc;
    

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, JSONException {
        
        
        //connect to local mongodb
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("TAS");
//        //user chooses file
        JFileChooser fileChooser = new JFileChooser();
        int returnValue = fileChooser.showOpenDialog(null);
//        //approve file chosen
        if(returnValue == JFileChooser.APPROVE_OPTION){
            
            
            //TRY CATCH
            //get selected file
            try {
                
                //read seleted excel file
                XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
                
                //get name, school, date
                Cell cID = wb.getSheetAt(0).getRow(11).getCell(1, xc.RETURN_BLANK_AS_NULL);
                String cid = new String();
                        cid = nullString(cID, cid);
                Cell cDate = wb.getSheetAt(0).getRow(8).getCell(1, xc.RETURN_BLANK_AS_NULL);
                String cdate = new String();
                        cdate = nullString(cDate, cdate);
                Cell cName = wb.getSheetAt(0).getRow(10).getCell(1, xc.RETURN_BLANK_AS_NULL);
                String cname = new String();
                        cname = nullString(cName, cname);
                Cell cSchool = wb.getSheetAt(0).getRow(12).getCell(1, xc.RETURN_BLANK_AS_NULL);
                String cschool = new String();
                        cschool = nullString(cSchool, cschool);
                
                        
                        //TEACHING
                Cell cCore = wb.getSheetAt(0).getRow(16).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double ccore = 0.0;
                        ccore = nullDouble(cCore, ccore);
                Cell cSupport = wb.getSheetAt(0).getRow(17).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double ctsupport = 0.0;
                        ctsupport = nullDouble(cSupport, ctsupport);
                
                        //RESEARCH
                Cell cCouncils = wb.getSheetAt(0).getRow(20).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double ccouncils = 0.0;
                        ccouncils = nullDouble(cCouncils, ccouncils);
                Cell cUK_govt = wb.getSheetAt(0).getRow(21).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cuk_govt = 0.0;
                        cuk_govt = nullDouble(cUK_govt, cuk_govt);
                Cell cEU = wb.getSheetAt(0).getRow(22).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double ceu = 0.0;
                        ceu = nullDouble(cEU, ceu);
                Cell cUK_charity = wb.getSheetAt(0).getRow(23).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cuk_charity = 0.0;
                        cuk_charity = nullDouble(cUK_charity, cuk_charity);
                Cell cUK_industry = wb.getSheetAt(0).getRow(24).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cuk_industry = 0.0; 
                        cuk_industry = nullDouble(cUK_industry, cuk_industry);
                Cell cKTP_projects = wb.getSheetAt(0).getRow(25).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cktp_projects = 0.0;
                        cktp_projects = nullDouble(cKTP_projects, cktp_projects);
                Cell cOther = wb.getSheetAt(0).getRow(26).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cother = 0.0;
                        cother = nullDouble(cOther, cother);
                Cell cSFC_innovation = wb.getSheetAt(0).getRow(27).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double csfc_innovation = 0.0;
                        csfc_innovation = nullDouble(cSFC_innovation, csfc_innovation);
                Cell cSFC_RD = wb.getSheetAt(0).getRow(28).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double csfc_rd = 0.0;
                        csfc_rd = nullDouble(cSFC_RD, csfc_rd);
                Cell cPGR_supervision = wb.getSheetAt(0).getRow(29).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cpgr_supervision = 0.0;
                        cpgr_supervision = nullDouble(cPGR_supervision, cpgr_supervision);
                Cell cInternal_research = wb.getSheetAt(0).getRow(30).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cinternal_research = 0.0;
                        cinternal_research = nullDouble(cInternal_research, cinternal_research);
                Cell cSupport_intext= wb.getSheetAt(0).getRow(31).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double csupport_intext = 0.0;
                        csupport_intext = nullDouble(cSupport_intext, csupport_intext);
                Cell cSupport_SFC = wb.getSheetAt(0).getRow(32).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double csupport_sfc = 0.0;
                       csupport_intext = nullDouble(cSupport_SFC, csupport_intext);
                
                       //SCHOLARSHIP
                Cell cTeaching = wb.getSheetAt(0).getRow(34).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cteaching = 0.0;
                    cteaching = nullDouble(cTeaching, cteaching);
                Cell cResearch = wb.getSheetAt(0).getRow(35).getCell(2, xc.RETURN_BLANK_AS_NULL);
               Double cresearch = 0.0;
                        cresearch = nullDouble(cResearch, cresearch);
                Cell cPhD = wb.getSheetAt(0).getRow(36).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cphd = 0.0;
                        cphd = nullDouble(cPhD, cphd);
                
                        //OTHER
                Cell coOther = wb.getSheetAt(0).getRow(38).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double coother = 0.0;
                        coother = nullDouble(coOther, coother);
               Cell coSupport = wb.getSheetAt(0).getRow(39).getCell(2, xc.RETURN_BLANK_AS_NULL);
               Double cosupport = 0.0;
                        cosupport = nullDouble(coSupport, cosupport);
                
                        //MANAGEMENT
                Cell cMgmt = wb.getSheetAt(0).getRow(41).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cmgmt = 0.0;
                        cmgmt = nullDouble(cMgmt, cmgmt);
                
                        //TOTAL
                Cell cTotal = wb.getSheetAt(0).getRow(43).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double ctotal = 0.0;
                      //  ctotal = nullDouble(cTotal, ctotal);
                
                      //HOLIDAYS
                Cell cHols = wb.getSheetAt(0).getRow(45).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double chols = 0.0;
                        chols = nullDouble(cHols, chols);
                
                        //MongoDB database object
                BasicDBObject document = new BasicDBObject();
                document.put("uID", cid);
                document.put("date", cdate);
                document.put("name", cname);
                document.put("school", cschool);
                
                //nested TEACHING document
                BasicDBObject documentTeach = new BasicDBObject();
                documentTeach.put("core", ccore);
                documentTeach.put("support", ctsupport);
                document.put("Teaching", documentTeach);
                
                //nested RESEARCH document
                BasicDBObject documentResearch = new BasicDBObject();
                documentResearch.put("council", ccouncils);
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
                
                //nested SCHOLARSHIP document
                BasicDBObject documentSchol = new BasicDBObject();
                documentSchol.put("teaching", cteaching);
                documentSchol.put("research", cresearch);
                documentSchol.put("PhD", cphd);
                document.put("Scholarship", documentSchol);
                
                //nested OTHER document
                BasicDBObject documentOther = new BasicDBObject();
                documentOther.put("Other", coother);
                documentOther.put("Osupport", cosupport);
                document.put("Other", documentOther);
                
                //add all to same mongodb document
                document.put("Mgmt", cmgmt);
                document.put("Total", ctotal);
                document.put("Hols", chols);
               
                
                //cursor
                //search mongodb TAS for document == cid
                //if doesn't exist, create
                //if does exist, update
                //collection.insert(document);
                
                BasicDBObject query = new BasicDBObject("uID", cid);
                
                
                
                }
            
                catch (FileNotFoundException ex) {
                    ex.printStackTrace(); 
                }
        
        
            }
        }
    
}
