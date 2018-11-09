/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tas;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.Mongo;
import java.io.File;
import java.io.FileFilter;
import java.io.IOException;
import java.net.UnknownHostException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.Date;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.joda.time.*;

/**
 *
 * @author fl8328
 */
public class GetHols {
    //collection dates
    //1/10 - 31/01
    //1/2-31/5
    //1/6 - 30/9
    static Double sem1 = 0.0;
    static Double sem2 = 0.0;
    static Double sem3 = 0.0;
    
    
     //Check if cell is null deals with NullPointerException
    //if DoubleCell == blank, d = 0.0
    public static Double nullDouble(Cell c, Double d){
        if(c == null){
            d = 0.0;
        }
        //if DoubleCell != blank, d = value of cell
        else{
            String con = c.toString();
            d = Double.parseDouble(con);
        }
        return d;
    }
    
    //Check if cell is null deals with NullPointerException
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
    
    int year = Calendar.getInstance().get(Calendar.YEAR);
   
    public static LocalDate nullDate(Cell c, LocalDate d){
        int month = 0;
       
       SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
       dateFormat.format(c.getDateCellValue());
                if(c == null){
                    d = LocalDate.parse("1970-01-01");
                }
                else{
                    d = LocalDate.parse(c.toString());
                    System.out.println(d);
                }
        return d;
    }
    
    public static int periodChecker(LocalDate datetime){
        int period = 0;
        
        //FIRST SEMESTER CUT OFF DATES
        if(datetime.getYear() == 1970){
            period = 0;
        }
        else if(datetime.getMonthOfYear() >= 10){
            period = 1;
        }
        else if(datetime.getMonthOfYear() == 1){
            period = 1;
        }
        
        
        //SECOND SEMESTER CUT OFF DATES
        else if(datetime.getMonthOfYear() == 2){
            period = 2;
        }
        else if(datetime.getMonthOfYear() == 3){
            period = 2;
        }
        else if(datetime.getMonthOfYear() == 4){
            period = 2;
        }
        if(datetime.getMonthOfYear() == 5){
            period = 2;
        }
        
        
        //THIRD SEMESTER CUT OFF DATES
        if(datetime.getMonthOfYear() == 6){
            period = 3;
        }
        if(datetime.getMonthOfYear() == 7){
            period = 3;
        }
        if(datetime.getMonthOfYear() == 8){
            period = 3;
        }
        if(datetime.getMonthOfYear() == 9){
            period = 3;
        }
        //return period
        return period;
    }
    
    //converts blank cell to null
    private static Row.MissingCellPolicy xc;
    
    public static void getHols() throws IOException, InvalidFormatException, ParseException{
        
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("TAS");
        
        final File folder = new File("H:\\NetBeansProjects\\TAS\\hol\\Ahriz, Hatem.xlsx");
        FileFilter filter = new ExcelFileFilter();
        File[] files = folder.listFiles(filter);
        
        //read excel file
        XSSFWorkbook wb = new XSSFWorkbook(folder);
        BasicDBObject document = new BasicDBObject();
        Cell name = wb.getSheetAt(0).getRow(4).getCell(1, xc.RETURN_BLANK_AS_NULL);
        String cname = new String();
        cname = nullString(name, cname);
        cname += "hols";
        Cell ent = wb.getSheetAt(0).getRow(6).getCell(1, xc.RETURN_BLANK_AS_NULL);
        Double cent = 0.0;
        cent = nullDouble(ent, cent);
        Cell carried = wb.getSheetAt(0).getRow(8).getCell(1, xc.RETURN_BLANK_AS_NULL);
        Double ccarried = 0.0;
        ccarried = nullDouble(carried, ccarried);
        Double total = cent + ccarried;
        for(int i = 18; i < 33; i++){
            Cell req = wb.getSheetAt(0).getRow(i).getCell(0, xc.RETURN_BLANK_AS_NULL);
           // System.out.println(req.toString());
            LocalDate creq = new LocalDate();
            creq = nullDate(req, creq);
            int sem = periodChecker(creq);
            //System.out.println(sem);
            Cell cdays = wb.getSheetAt(0).getRow(i).getCell(2, xc.RETURN_BLANK_AS_NULL);
//            System.out.println(cdays.toString());
            if(sem == 1){
                sem1 += nullDouble(cdays, sem1);
                System.out.println("Semester 1");
                document.put("name", cname);
                document.put("sem1", sem1);
            }
            else if(sem == 2){
                sem2 += nullDouble(cdays, sem2);
                System.out.println("Semester 2");
                document.put("name", cname);
                document.put("sem2", sem2);
            }
            else if(sem == 3){
                sem3 += nullDouble(cdays, sem3);
                System.out.println("Semester 3");
                document.put("name", cname);
                document.put("sem3", sem3);
            }
            else{
            System.out.println("empty");
            }
        }
        BasicDBObject query = new BasicDBObject("name", cname);
        //if user ID is already in the DB then update the entry
        //if user ID is not in the DB then add them
        DBCursor cursor = collection.find(query);
        if(cursor.hasNext()){
            collection.update(cursor.next(), document);
            System.out.println("Updated document " + cname );
        }
        else{
            collection.insert(document);
            System.out.println("Added document " + cname);
        }
    }
}

