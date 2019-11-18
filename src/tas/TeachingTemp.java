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
public class TeachingTemp {
    static int p = 5;
    
    
      public static float nullFloat(Cell c, float fl){
        if(c == null){
            fl = 0.00000f;
        }
        //if DoubleCell != blank, d = value of cell
        else{
            String con = c.toString();
        }
        return fl;
    }
    
     public static Date nullDate(Cell c, String d) throws ParseException{

        Date date = new Date();
       //date format is <Day>.<Month>.<Year>
       SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
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
    
    public static void teachingTemp() throws IOException, ParseException{
        //connect to local mongodb
        Mongo mongo = new Mongo("localhost", 27017);
        DB tas = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = tas.getCollection("NEWTASTEMP");
        
        //remove all entries from the database
        DBCursor rem = collection.find();
        while(rem.hasNext()){
            collection.remove(rem.next());
        }
        //user chooses directory containing all users tas excel sheets
        
         //read excel file
                XSSFWorkbook wb = new XSSFWorkbook("Copy of TAS-Workload Model_18-19v2.xlsx");                             
                
                for(int i = 5; i <= 32; i++){
                    
                    BasicDBObject document = new BasicDBObject();
                        
                        Cell cSur = wb.getSheetAt(0).getRow(i).getCell(0, xc.CREATE_NULL_AS_BLANK);
                        String sur = cSur.getStringCellValue();
                        
                        Cell cFir = wb.getSheetAt(0).getRow(i).getCell(1, xc.CREATE_NULL_AS_BLANK);
                        String fir = cFir.getStringCellValue();
                        document.put("name", fir + " " + sur);
                        
                        Cell cID = wb.getSheetAt(0).getRow(i).getCell(2, xc.CREATE_NULL_AS_BLANK);
                        String id = cID.getStringCellValue();
                        id = id.toUpperCase();
                        document.put("id", id);
                        
                       Cell tTot = wb.getSheetAt(0).getRow(i).getCell (12, xc.CREATE_NULL_AS_BLANK);
                       Double tot = (tTot.getNumericCellValue()*p)/100;
                       document.put("tot", tot);
                        
                        Cell cR = wb.getSheetAt(0).getRow(i).getCell(13, xc.CREATE_NULL_AS_BLANK);
                        Double r = (cR.getNumericCellValue()*p)/100;
                        document.put("research", r);
                        
                        
                        Cell cCother = wb.getSheetAt(0).getRow(i).getCell(14, xc.CREATE_NULL_AS_BLANK);
                        Double cother = (cCother.getNumericCellValue()*p)/100;
                        document.put("c other", cother);           
                        
                        Cell cCotherSupp = wb.getSheetAt(0).getRow(i).getCell(15, xc.CREATE_NULL_AS_BLANK);
                        Double cOtherSupp = (cCotherSupp.getNumericCellValue()*p)/100;
                        document.put("cother supp", cOtherSupp);
                                  
                        Cell cOtherT = wb.getSheetAt(0).getRow(i).getCell(16, xc.CREATE_NULL_AS_BLANK);
                        Double otherT = (cOtherT.getNumericCellValue()*p)/100;
                        document.put("otherT", otherT);
                        
                        
                        Cell cOtherpgr = wb.getSheetAt(0).getRow(i).getCell(17, xc.CREATE_NULL_AS_BLANK);
                        Double otherpgr = (cOtherpgr.getNumericCellValue()*p)/100;
                        document.put("pgr", otherpgr);
                        
                        
                        Cell cOthersupp = wb.getSheetAt(0).getRow(i).getCell(18, xc.CREATE_NULL_AS_BLANK);
                        Double othersupp = (cOthersupp.getNumericCellValue()*p)/100;
                        document.put("other supp", othersupp);
                        
                        
                        Cell cOtherfurther = wb.getSheetAt(0).getRow(i).getCell(19, xc.CREATE_NULL_AS_BLANK);
                        Double otherfurther = (cOtherfurther.getNumericCellValue()*p)/100;
                        document.put("further", otherfurther);
                        
                        Cell cOthermgmt = wb.getSheetAt(0).getRow(i).getCell(20, xc.CREATE_NULL_AS_BLANK);
                        Double othermgmt = (cOthermgmt.getNumericCellValue()*p)/100;
                        document.put("mgmt", othermgmt);
                        
                        
                        Cell ctime = wb.getSheetAt(0).getRow(i).getCell(23, xc.CREATE_NULL_AS_BLANK);
                        Double time = ctime.getNumericCellValue();
                        document.put("RemTime", time);
                        
                        document.put("sem1", 0.0);
                        document.put("sem2", 0.0);
                        document.put("sem3", 0.0);
                        
                        collection.insert(document);
                        
                        
                        System.out.println("------------------------------------------------");
                        System.out.println(sur + "\n" + fir + "\n" + tot + "\n" + r + "\n" + cother + "\n" + cOtherSupp + "\n" + otherT + "\n" + otherpgr + "\n" + othersupp + "\n" + otherfurther + "\n" + othermgmt + "\n" + time);      
                               
                    
                }
                
    }
                
}
