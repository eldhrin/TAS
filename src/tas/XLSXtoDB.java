// Adam Lyons 20/09/2018
//This class gets a directory chosen by the user which will loop through 
//and read all .xls or .xlsx files (legacy and modern excel files)
//This is read and then added to a local mongoDB
//if the document already exists (determined by the person's login ID)
//the database updates that document
//if the document does not already exist, the document is created

//if there is a DB error, the database is offline and needs reconnected

package tas;

//import files
import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.Mongo;
import java.io.File;
import java.io.FileFilter;
import java.io.FileNotFoundException;
import java.io.IOException;
import javax.swing.JFileChooser;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//Start of class GetXlsx
public class XLSXtoDB{
    
    //converts blank cell to null
    private static Row.MissingCellPolicy xc;
    
    public static void xlsxtoDB() throws IOException, InvalidFormatException{
        //connect to local mongodb
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("TAS");
        //user chooses directory containing all users tas excel sheets
        
        DBCursor rem = collection.find();
        while(rem.hasNext()){
            collection.remove(rem.next());
        }
        
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new java.io.File("."));
        chooser.setDialogTitle("Choose a directory");
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.setAcceptAllFileFilterUsed(false);
        
        if(chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION){
        
        FileFilter filter = new ExcelFileFilter();
        File directory = chooser.getSelectedFile();
        File[] files = directory.listFiles(filter);
        //for every file in the chosen directory
        for(File file : files){
            
            
            //TRY CATCH
            try {
                
                //read excel file
                XSSFWorkbook wb = new XSSFWorkbook(file);
                System.out.println(file);
                
                
                //GET VARIABLES FROM THE SPREADSHEET AND CONVERT TO STRING/DOUBLE
                //get name, school, date
                Cell cID = wb.getSheetAt(0).getRow(11).getCell(1, xc.RETURN_BLANK_AS_NULL);
                String cid = new String();
                        cid = Null.nullString(cID, cid).toString();
                if(cid.equals("")){
                    
                }
                Cell cDate = wb.getSheetAt(0).getRow(8).getCell(1, xc.RETURN_BLANK_AS_NULL);
                String cdate = new String();
                        cdate = Null.nullString(cDate, cdate);
                Cell cName = wb.getSheetAt(0).getRow(10).getCell(1, xc.RETURN_BLANK_AS_NULL);
                String cname = new String();
                        cname = Null.nullString(cName, cname);
                Cell cSchool = wb.getSheetAt(0).getRow(12).getCell(1, xc.RETURN_BLANK_AS_NULL);
                String cschool = new String();
                        cschool = Null.nullString(cSchool, cschool);
                      
                //TEACHING
                Cell cCore = wb.getSheetAt(0).getRow(16).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double ccore = 0.0;
                        ccore = Null.nullDouble(cCore, ccore);
                Cell cSupport = wb.getSheetAt(0).getRow(17).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double ctsupport = 0.0;
                        ctsupport = Null.nullDouble(cSupport, ctsupport);
                
                //RESEARCH
                Cell cCouncils = wb.getSheetAt(0).getRow(20).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double ccouncils = 0.0;
                        ccouncils = Null.nullDouble(cCouncils, ccouncils);
                Cell cUK_govt = wb.getSheetAt(0).getRow(21).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cuk_govt = 0.0;
                       cuk_govt = Null.nullDouble(cUK_govt, cuk_govt);
                Cell cEU = wb.getSheetAt(0).getRow(22).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double ceu = 0.0;
                        ceu = Null.nullDouble(cEU, ceu);
                Cell cUK_charity = wb.getSheetAt(0).getRow(23).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cuk_charity = 0.0;
                        cuk_charity = Null.nullDouble(cUK_charity, cuk_charity);
                Cell cUK_industry = wb.getSheetAt(0).getRow(24).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cuk_industry = 0.0; 
                        cuk_industry = Null.nullDouble(cUK_industry, cuk_industry);
                Cell cKTP_projects = wb.getSheetAt(0).getRow(25).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cktp_projects = 0.0;
                        cktp_projects = Null.nullDouble(cKTP_projects, cktp_projects);
                Cell cOther = wb.getSheetAt(0).getRow(26).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cother = 0.0;
                        cother = Null.nullDouble(cOther, cother);
                Cell cSFC_innovation = wb.getSheetAt(0).getRow(27).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double csfc_innovation = 0.0;
                        csfc_innovation = Null.nullDouble(cSFC_innovation, csfc_innovation);
                Cell cSFC_RD = wb.getSheetAt(0).getRow(28).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double csfc_rd = 0.0;
                        csfc_rd = Null.nullDouble(cSFC_RD, csfc_rd);
                Cell cPGR_supervision = wb.getSheetAt(0).getRow(29).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cpgr_supervision = 0.0;
                        cpgr_supervision = Null.nullDouble(cPGR_supervision, cpgr_supervision);
                Cell cInternal_research = wb.getSheetAt(0).getRow(30).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cinternal_research = 0.0;
                        cinternal_research = Null.nullDouble(cInternal_research, cinternal_research);
                Cell cSupport_intext= wb.getSheetAt(0).getRow(31).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double csupport_intext = 0.0;
                        csupport_intext = Null.nullDouble(cSupport_intext, csupport_intext);
                Cell cSupport_SFC = wb.getSheetAt(0).getRow(32).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double csupport_sfc = 0.0;
                       csupport_sfc = Null.nullDouble(cSupport_SFC, csupport_sfc);
                
                //SCHOLARSHIP
                Cell cTeaching = wb.getSheetAt(0).getRow(34).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cteaching = 0.0;
                   cteaching = Null.nullDouble(cTeaching, cteaching);
                Cell cResearch = wb.getSheetAt(0).getRow(35).getCell(2, xc.RETURN_BLANK_AS_NULL);
               Double cresearch = 0.0;
                        cresearch = Null.nullDouble(cResearch, cresearch);
                Cell cPhD = wb.getSheetAt(0).getRow(36).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double cphd = 0.0;
                        cphd = Null.nullDouble(cPhD, cphd);
              
              //OTHER
              Cell coOther = wb.getSheetAt(0).getRow(38).getCell(2, xc.RETURN_BLANK_AS_NULL);
              Double coother = 0.0;
                      coother = Null.nullDouble(coOther, coother);
               Cell coSupport = wb.getSheetAt(0).getRow(39).getCell(2, xc.RETURN_BLANK_AS_NULL);
               Double cosupport = 0.0;
                        cosupport = Null.nullDouble(coSupport, cosupport);
              
              //MANAGEMENT
              Cell cMgmt = wb.getSheetAt(0).getRow(41).getCell(2, xc.RETURN_BLANK_AS_NULL);
              Double cmgmt = 0.0;
                      cmgmt = Null.nullDouble(cMgmt, cmgmt);
                
                //TOTAL
                Double ctotal = ccore + ctsupport + ccouncils  + cuk_govt + ceu + cuk_charity + cuk_industry + cktp_projects + cother + csfc_innovation + csfc_rd + cpgr_supervision + cinternal_research + csupport_intext + csupport_sfc + cteaching + cresearch + cphd + coother + cosupport + cmgmt;
                
                //HOLIDAYS
                Cell cHols = wb.getSheetAt(0).getRow(45).getCell(2, xc.RETURN_BLANK_AS_NULL);
                Double chols = 0.0;
                        chols = Null.nullDouble(cHols, chols);
                
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
                
//                //add all to same mongodb document
                document.put("Mgmt", cmgmt);
                document.put("Total", ctotal);
                document.put("Hols", chols);

                collection.insert(document);
                System.out.println("Added document " + cname + " with ID " + cid);
                //close the database
                
                }
                catch (FileNotFoundException ex) {
                    ex.printStackTrace(); 
                }
            System.out.println("Excel file written to the database");
        
        
            }//end of directory loop
        mongo.close();
        }//end of directory chooser
    }//end of getXlsx()
}// end of main class
