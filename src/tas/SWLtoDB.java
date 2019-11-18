//Adam Lyons 20/11/2018
//STAFF WORKLOAD XLSX TO DATABASE
//USES ID FOR UNIQUE ID
package tas;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.Mongo;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.JOptionPane;
import org.apache.poi.xssf.usermodel.XSSFSheet;


public class SWLtoDB {
   static int period;
    
    public static void SWLtoDB() throws IOException, InvalidFormatException{
        
        
        int  d = Null.getDate();
        String time = "";
        Date y = new Date();
        Calendar cal = Calendar.getInstance();
        cal.setTime(y);
        int year = cal.get(Calendar.YEAR);
       //if semester == <number> then writes the appropriate collection period to the cell
       switch (d) {
           case 1:
               time = "1st of June " + cal.get(Calendar.YEAR) + " to 31st of August " + year;
               break;
           case 2:
               time = "1st of October " + (year-1) + " to 31st of January " + year;
               break;
           default:
               time = "1st of Feburary " + cal.get(Calendar.YEAR) + " to 31st of May " + year;
               break;
       }
        
        //connect to DB
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("NEWTAS");
        
        //Clear the database before writing to it again, prevents old members of staff being 'stuck' in the 
        //database with no way to remove    
        DBCursor rem = collection.find();
        while(rem.hasNext()){
            collection.remove(rem.next());
        }
        //gets how many entries are in the DB (starts from 1)
        int dbcount = (int)collection.count();
        //uses this file as a base to write to
        FileInputStream file = new FileInputStream(new File("Staff Workload Model-v2.xlsx"));
        XSSFWorkbook wb = new XSSFWorkbook(file);
        XSSFSheet sheet = wb.getSheetAt(0);
        
        FileOutputStream fileOut = new FileOutputStream("S:\\Computing\\TAS\\TAS-Workload Model_18-19.xlsx");
            
        for(int i = 2; i <= 31; i++){
            try{
                Cell cfName = wb.getSheetAt(0).getRow(i).getCell(1);
                String fName = "";
                fName = Null.nullString(cfName, fName);

                Cell clName = wb.getSheetAt(0).getRow(i).getCell(0);
                String lName = "";
                lName = Null.nullString(clName, lName);
                String fullName = fName + " " + lName;
                System.out.println(fullName);
                
                    Cell cuid = wb.getSheetAt(0).getRow(i).getCell(2);
                    String uid = new String();
                    uid = Null.nullString(cuid, uid);
                    System.out.println(uid);
                    
                    Cell cTeaching = wb.getSheetAt(0).getRow(i).getCell(3);
                    Double teaching = 0.0;
                    teaching = Null.nullDouble(cTeaching, teaching);
                    Double teachingp = (teaching * 5)/100;
                    System.out.println("teaching " + teachingp);
                    
                    Cell cTSupport = wb.getSheetAt(0).getRow(i).getCell(4);
                    Double tSupport = 0.0;
                    tSupport = Null.nullDouble(cTSupport, tSupport);
                    Double tSupportp = (tSupport * 5)/100;
                    System.out.println("teaching support " + tSupportp);
                    
                    Cell cga = wb.getSheetAt(0).getRow(i).getCell(5);
                    Double ga = 0.0;
                    ga = Null.nullDouble(cga, ga);
                    Double gap = (ga*5)/100;
                    System.out.println("ga " + gap);
                    
                    Cell cservice = wb.getSheetAt(0).getRow(i).getCell(6);
                    Double service = 0.0;
                    service = Null.nullDouble(cservice, service);
                    Double servicep = (service*5)/100;
                    System.out.println("service courses " + servicep);
                    
                    Cell cResearch = wb.getSheetAt(0).getRow(i).getCell(7);
                    Double tResearch = 0.0;
                    tResearch = Null.nullDouble(cResearch, tSupport);
                    Double tResearchp = (tResearch * 5)/100;
                    System.out.println("research " + tResearchp);
                    
                    Cell ccomm = wb.getSheetAt(0).getRow(i).getCell(8);
                    Double comm = 0.0;
                    comm = Null.nullDouble(ccomm, comm);
                    Double commp = (comm*5)/100;
                    System.out.println("commercial " + commp);
                    
                    Cell cOther = wb.getSheetAt(0).getRow(i).getCell(9);
                    Double other = 0.0;
                    other = Null.nullDouble(cOther, other);
                    Double otherp = (other*5)/100;
                    System.out.println("other" + otherp);
                    
                                    //MongoDB database object
                //create DB Object and add name, date, school
                BasicDBObject document = new BasicDBObject();
                document.put("uID", uid);
                document.put("date", time);
                document.put("name", fullName);
                document.put("school", "CSDM");
                    
                BasicDBObject documentTeach = new BasicDBObject();
                documentTeach.put("core", teachingp + gap);
                documentTeach.put("support", tSupportp);
                //insert into main document
                document.put("Teaching", documentTeach);
                
                //nested RESEARCH document
                //new DB Object nested in main, adds research section
                BasicDBObject documentResearch = new BasicDBObject();
                documentResearch.put("council", tResearchp);
//                documentResearch.put("UK_govt", cuk_govt);
//                documentResearch.put("EU", ceu);
//                documentResearch.put("UK_charity", cuk_charity);
//                documentResearch.put("UK_industry", cuk_industry);
//                documentResearch.put("KTP_projects", cktp_projects);
//                documentResearch.put("other", cother);
//                documentResearch.put("SFC_innovation", csfc_innovation);
//                documentResearch.put("SFC_RD", csfc_rd);
//                documentResearch.put("PGR_supervision", cpgr_supervision);
//                documentResearch.put("internal_research", cinternal_research);
                //documentResearch.put("support_intext", commp);
//                documentResearch.put("support_SFC", csupport_sfc);
                //insert into main document
                document.put("Research", documentResearch);
                
                //nested SCHOLARSHIP document
                //new DB Object nested in main, adds scholarship section
                BasicDBObject documentSchol = new BasicDBObject();
//                documentSchol.put("teaching", cteaching);
//                documentSchol.put("research", cresearch);
//                documentSchol.put("PhD", cphd);
                //insert into main document
//                document.put("Scholarship", documentSchol);
                
                //nested OTHER document
                //new DB Object nested in main, adds other section
                BasicDBObject documentOther = new BasicDBObject();
                documentOther.put("Other", commp);
//                documentOther.put("Osupport", cosupport);
                //insert it into main document
                document.put("Other", documentOther);
                
                Double e = 0.0;
                
                //add to main DB Object
                document.put("Mgmt", otherp);
//                document.put("Total", ctotal);
//                document.put("Hols", chols);
                document.put("sem1", e);
                document.put("sem2", e);
                document.put("sem3", e);
                collection.insert(document);
                

            }
            catch (Exception e){
                e.printStackTrace();
                System.out.println("Database might not be initalised");
            }
        }
        System.out.println("Workload excel file to the database");
            wb.write(fileOut);
       //close workbook
        wb.close();
        //close DB connection
        mongo.close();
        JOptionPane.showMessageDialog(null, "Done! Written to Database");
    }//end generateXlsx()
}//end class
