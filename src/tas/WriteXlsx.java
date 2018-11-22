//Adam Lyons 03/10/2018
//This class generates the main TAS report with all the users who hae submitted a TAS return form

package tas;

import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.Mongo;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Calendar;
import java.util.Date;
import java.util.Scanner;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author fl8328
 */
public class WriteXlsx {
    static int period;
    
    //function gets the semester based on today's date
    public static int getDate(){
        //collection dates
        //1/10 - 31/01
        //1/2-31/5
        //1/6 - 30/9
        Date sem = new Date();
        Calendar cal = Calendar.getInstance();
        cal.setTime(sem);
        //Months of the year start from 0(being Janurary)
        //if this month is October, November, December then semester = 1
         if(cal.get(Calendar.MONTH) >= 9){
            period = 1;
        }
         //if this month is January, semester = 1
        else if(cal.get(Calendar.MONTH) == 0){
            period = 1;
        }
        
        
        //SECOND SEMESTER CUT OFF DATES
        //if this month is Feburary, March, April, May then semester = 2
        else if(cal.get(Calendar.MONTH) == 1){
            period = 2;
        }
        else if(cal.get(Calendar.MONTH) == 2){
            period = 2;
        }
        else if(cal.get(Calendar.MONTH) == 3){
            period = 2;
        }
        if(cal.get(Calendar.MONTH) == 4){
            period = 2;
        }
        
        
        //THIRD SEMESTER CUT OFF DATES
        //if this month is June, July, August, September then semester = 3
        if(cal.get(Calendar.MONTH) == 5){
            period = 3;
        }
        if(cal.get(Calendar.MONTH) == 6){
            period = 3;
        }
        if(cal.get(Calendar.MONTH) == 7){
            period = 3;
        }
        if(cal.get(Calendar.MONTH) == 8){
            period = 3;
        }
        //return period
        return period;
        
    }//end getDate()
    
    
    public static XSSFWorkbook writeXlsx() throws IOException{
        
        //connect to the database
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //count is first row where we write the data to
        int count = 9;
        JFrame frame = new JFrame("Input");
        String in = JOptionPane.showInputDialog(frame, "How many academic staff do we currently have?");
        int n = Integer.parseInt(in);
        //find collection TAS
        DBCollection collection = db.getCollection("TAS");
        DBCursor cursor = collection.find();  
        //get number of entries in the database (this starts at 1)
        int dbcount = (int)collection.count();
        XSSFWorkbook wb = new XSSFWorkbook();
        //base file the program writes to
        FileOutputStream fileOut = new FileOutputStream("megareportKEEP.xlsx");
        //file that the program outputs (creates this file if it does not exist)
        String pathName = "H:\\NetBeansProjects\\TAS\\workbook.xlsx";
        wb = new XSSFWorkbook(pathName);
                     
        //INIT
        //initialises the file
        //writes the collection period based on today's date
        Cell tdate = wb.getSheetAt(0).getRow(3).getCell(2);
        int  d = getDate();
        String time = "";
        Date y = new Date();
        Calendar cal = Calendar.getInstance();
        cal.setTime(y);
        //if semester == <number> then writes the appropriate collection period to the cell
        if(d == 1){
            time = "1st of October " + cal.get(Calendar.YEAR) + " to 31st of Januaray " + cal.get(Calendar.YEAR+1);
        }
        else if (d == 2){
            time = "1st of October " + cal.get(Calendar.YEAR) + " to 31st of Januaray " + cal.get(Calendar.YEAR);
        }
        else{
            time = "1st of October " + cal.get(Calendar.YEAR) + " to 31st of Januaray " + cal.get(Calendar.YEAR);
        }
        tdate.setCellValue(time);
                    
        Cell tschool = wb.getSheetAt(0).getRow(5).getCell(2);
        tschool.setCellValue("CSDM");
                    
        //loops through for every entry in database
        for(int j = 0; j < dbcount; j++){
                //gets the data from the current database entry its looking at
                DBObject o = cursor.next();
                DBObject t = (DBObject) o.get("Teaching");
                DBObject r = (DBObject) o.get("Research");
                DBObject s = (DBObject) o.get("Scholarship");
                DBObject q = (DBObject) o.get("Other");
                
                    
                //USER
                //gets name of person
                    Cell tname = wb.getSheetAt(0).getRow(count).getCell(0);
                    tname.setCellValue(o.get("name").toString());
                //gets ID
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
                   
                    if(d == 1){               
                    Double hols = Double.parseDouble(o.get("sem1").toString());
                    thols.setCellValue(hols);
                    }
                    //if today's date is semester 2, get semester 2's holidays
                    else if(d == 2){          
                    Double hols = Double.parseDouble(o.get("sem2").toString());
                    thols.setCellValue(hols);
                    }
                    //if today's date is semester 3, get semester 3's holidays
                    else{              
                    Double hols = Double.parseDouble(o.get("sem3").toString());
                    thols.setCellValue(hols);
                    }
                    
                
                //calc TAS RESEARCH
                //TAS Research is seperated at the end
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

                    //total % TAS research
                    Double tasresearch = council + UK_govt + EU + UK_charity + UK_industry + KTP_projects + R_other + SFC_innovation + SFC_RD + internal_research + support_intext + support_SFC/100;
                    Cell tas = wb.getSheetAt(0).getRow(count).getCell(26);
                    tas.setCellValue(tasresearch);

                    //++ row that the DB is writing to
                    count++;
            }//ends for
             
            wb.write(fileOut);
            fileOut.close();
            //close workbook
            wb.close();
            
            //calculates if we have reached 85% return (based on number given)
            //needs 85% as set by HR
            int t = (int)Math.ceil((n/100)*85);
            //if less than 85% then warning message 
            if(dbcount < t){
                JOptionPane.showMessageDialog(null, "You have less than 85%", "Warning: " + "Info", JOptionPane.INFORMATION_MESSAGE);
            }
            //else "Done!" message
            else{
                JOptionPane.showMessageDialog(null, "Done! \n Report is saved as " + pathName, "Info", JOptionPane.INFORMATION_MESSAGE);
            }
        //close mongodb
        mongo.close();
        //highlighting missing return statement, however none is needed but it errors without it
        return null;
    }//end WriteXlsx()
        
}//end class
