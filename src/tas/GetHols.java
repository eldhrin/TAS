//Adam Lyons 20/11/2018
//This class gets the holidays from a seperate file and writes that to the database
//of the corresponding member of staff
package tas;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.Mongo;
import java.io.File;
import java.io.FileFilter;
import java.io.IOException;
import java.text.ParseException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.Date;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

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
    static int i = 18;
  
    //converts blank cell to null
    private static Row.MissingCellPolicy xc;
    
    public static void getHols() throws IOException, InvalidFormatException, ParseException{
        //connect to database
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("NEWTASTEMP");
        //user chooses directory containing all users tas excel sheets
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
                BasicDBObject document = new BasicDBObject();
                //gets name
                Cell cuid = wb.getSheetAt(0).getRow(5).getCell(1, xc.RETURN_BLANK_AS_NULL);
                String id = new String();
                id = Null.nullString(cuid, id);
                //finds the database object with the name given
                //ID MUST BE PRESENT IN FIELD OR ELSE IT WON'T WRITE TO DATABASE PROPERLY
                BasicDBObject query = new BasicDBObject("id", id);
                System.out.println(id);
                DBCursor cursor = collection.find(query);
                
                //if the object is found then add to it
                if(cursor.hasNext()){
                    System.out.println(id);
                    System.out.println("FOUND ID " + id);
                    //holiday entitlement
                    Cell ent = wb.getSheetAt(0).getRow(6).getCell(1, xc.RETURN_BLANK_AS_NULL);
                    Double cent = 0.0;
                    cent = Null.nullDouble(ent, cent);
                    //holidays carried from last year
                    Cell carried = wb.getSheetAt(0).getRow(8).getCell(1, xc.RETURN_BLANK_AS_NULL);
                    Double ccarried = 0.0;
                    ccarried = Null.nullDouble(carried, ccarried);
                    //total holidays
                    Double total = cent + ccarried;

                    //loop through the dates given (this starts at row 18)
                    for(int i = 18; i < 30; i++){
                        Cell req = wb.getSheetAt(0).getRow(i).getCell(0, xc.RETURN_BLANK_AS_NULL);
                        String creq = new String();
                        Date date = new Date();
                        date = Null.nullDate(req, creq);
                        //if cell is null then there are no more dates, skip to the next DB object
                        if(date == null){
                            if(sem1 != 0.0){
                                System.out.println(sem1);
                            }
                            else if(sem1 == null){
                                sem1 = 0.0;
                            }
                            else if(sem2 != 0.0){
                                System.out.println(sem2);
                            }
                            else if(sem2 == null){
                                sem2 = 0.0;
                            }
                            else if(sem3 != 0.0){
                                System.out.println(sem3);
                            }
                            else{
                                sem3 = 0.0;
                            }
                        }
                        //if the cell is not null then there is a date, process it
                        else{
                            //checks what semester the holiday falls in
                            int sem = Null.periodChecker(date);
                            //System.out.println(sem);
                            Cell cdays = wb.getSheetAt(0).getRow(i).getCell(2, xc.RETURN_BLANK_AS_NULL);
                            Double days = 0.0;
                            days = Null.nullDouble(cdays, days);
                            //if sem is 1 then add to current semester 1 holidays
                            if(sem == 1){
                                sem3 += days;
                            }
                            //if sem is 2 then add to current semester 2 holidays
                            else if(sem == 2){
                                sem1 += days;
                            }
                            //if sem is 3 then add to current semester 3 holidays
                            else if(sem == 3){
                                sem2 += days;
                            }
                        }//end if else
                    }//end for
                        //add to the associated database object
                        document.put("sem1", sem1);
                        System.out.println(sem1);
                        document.put("sem2", sem2);
                        System.out.println(sem2);
                        document.put("sem3", sem3);
                        System.out.println(sem3);
                        collection.update(cursor.next(), new BasicDBObject("$set",document));
                    }//end DB query loop
                    //if database object is not found then skip over the file
                    else {
                    System.out.println(id + " not present");
                    System.out.println("---------------------------");
                        continue;
                    }
                //reset the semester holiday variables
                sem1 = 0.0;
                sem2 = 0.0;
                sem3 = 0.0;
                System.out.println("-----------------------------");
            }
            catch (Exception e){
                e.printStackTrace();
                System.out.println("Database might not be initalised");
            }
        }//end for file loop
        System.out.println("Gotten holidays and added to db");
        JOptionPane.showMessageDialog(null, "Done! \n Saved to the database", "Info", JOptionPane.INFORMATION_MESSAGE);
    }//end getHols()
    }
}//end class

