// APPLICATIONS AND AWARDS FOR <YEAR> 
// PULLS DATA FROM XLSX AND ADDS TO DB


package tas;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
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
public class AwardstoDB {
    
     //checks if date cell is null or not
    public static Date nullDate(Cell c, String d) throws ParseException{

        Date date = new Date();
       //date format is <Day>.<Month>.<Year>
       SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                if(c == null){
                    date = null;
                }
                else{
                    date = dateFormat.parse(c.toString());
                }
        return date;
    }
    
    //converts blank cell to null
    private static Row.MissingCellPolicy xc;
    
    public static void awardstoDB() throws IOException, ParseException{
        //connect to local mongodb
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("TAS");
        //user chooses directory containing all users tas excel sheets
        
         //read excel file
                XSSFWorkbook wb = new XSSFWorkbook("ApplicationsAndAwards.xlsx");
                
                for(int i = 4; i < 23; i++){
                    Cell cProjectTitle = wb.getSheetAt(1).getRow(i).getCell(0, xc.CREATE_NULL_AS_BLANK);
                    String projectTitle = new String();
                    projectTitle = Null.nullString(cProjectTitle, projectTitle);
                    System.out.println(projectTitle);
                    
                    Cell cPID = wb.getSheetAt(1).getRow(i).getCell(1, xc.CREATE_NULL_AS_BLANK);
                    String PID =  new String();
                    PID = Null.nullString(cPID, PID);
                    System.out.println(PID);
                    
                    Cell cPLead = wb.getSheetAt(1).getRow(i).getCell(2, xc.CREATE_NULL_AS_BLANK);
                    String PLead = new String();
                    PLead = Null.nullString(cPLead, PLead);
                    System.out.println(PLead);
                    
                    Cell cco1 = wb.getSheetAt(1).getRow(i).getCell(3, xc.CREATE_NULL_AS_BLANK);
                    String co1 = new String();
                    co1 = Null.nullString(cco1, co1);
                    System.out.println(co1);
                    
                    Cell cco2 = wb.getSheetAt(1).getRow(i).getCell(4, xc.CREATE_NULL_AS_BLANK);
                    String co2 = new String();
                    co2 = Null.nullString(cco2, co2);
                    System.out.println(co2);
                    
                    Cell cco3 = wb.getSheetAt(1).getRow(i).getCell(5, xc.CREATE_NULL_AS_BLANK);
                    String co3 = new String();
                    co3 = Null.nullString(cco3, co3);
                    System.out.println(co3);
                    
                    Cell csDate = wb.getSheetAt(1).getRow(i).getCell(6, xc.CREATE_NULL_AS_BLANK);
                    String newDate = new String();
                    Date sDate = new Date();
                    sDate = nullDate(csDate, newDate);
                    System.out.println("Start date: " + sDate);
                    
                    Cell ceDate = wb.getSheetAt(1).getRow(i).getCell(7, xc.CREATE_NULL_AS_BLANK);
                    String secondDate = new String();
                    Date eDate = new Date();
                    eDate = nullDate(ceDate, secondDate);
                    System.out.println("end date: " + eDate);
                    System.out.println("--------------------------------------------------");
                   
                    BasicDBObject query = new BasicDBObject("name", PLead);
                    //if user ID is already in the DB then update the entry
                    //if user ID is not in the DB then add them
                    DBCursor cursor = collection.find(query);
                    BasicDBObject document = new BasicDBObject();
                    if(cursor.hasNext()){
                        document.put("project", projectTitle);
                        collection.update(cursor.next(), new BasicDBObject("$set",document));
                    }
                    
                    BasicDBObject query01 = new BasicDBObject("name", co1);
                    DBCursor cursor01 = collection.find(query01);
                    if(cursor01.hasNext()){
                        document.put("project1", projectTitle);
                        collection.update(cursor01.next(), new BasicDBObject("$set",document));
                    }
                    
                    BasicDBObject query02 = new BasicDBObject("name", co2);
                    DBCursor cursor02 = collection.find(query02);
                    if(cursor02.hasNext()){
                        document.put("project2", projectTitle);
                        collection.update(cursor02.next(), new BasicDBObject("$set",document));
                    }
                    
                    BasicDBObject query03 = new BasicDBObject("name", co3);
                    DBCursor cursor03 = collection.find(query03);
                    if(cursor03.hasNext()){
                        document.put("project3", projectTitle);
                        collection.update(cursor03.next(), new BasicDBObject("$set",document));
                    }
                }
                    
            }
               
    }
