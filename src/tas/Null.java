//NULL DEALS WITH METHODS THAT DEAL WITH NULL VALUES AND DATES
package tas;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;

/**
 *
 * @author fl8328
 */
public class Null {
     static int period;
    
   
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
    
   
    //function gets the semester based on today's date
    public static int getDate(){
        //collection dates
        //1/10 - 31/01
        //1/2 - 31/5
        //1/6 - 30/9
        Date sem = new Date();
        Calendar cal = Calendar.getInstance();
        cal.setTime(sem);
        //Months of the year start from 0(being Janurary)
        //if this month is October, November, December then semester = 1
         if(cal.get(Calendar.MONTH) >= 9){
            period = 2;
        }
         //if this month is January, semester = 1
        else if(cal.get(Calendar.MONTH) == 0){
            period = 2;
        }
        
        
        //SECOND SEMESTER CUT OFF DATES
        //if this month is Feburary, March, April, May then semester = 2
        else if(cal.get(Calendar.MONTH) == 1){
            period = 3;
        }
        else if(cal.get(Calendar.MONTH) == 2){
            period = 3;
        }
        else if(cal.get(Calendar.MONTH) == 3){
            period = 3;
        }
        else if(cal.get(Calendar.MONTH) == 4){
            period = 3;
        }
        
        
        //THIRD SEMESTER CUT OFF DATES
        //if this month is June, July, August, September then semester = 3
        else if(cal.get(Calendar.MONTH) == 5){
            period = 1;
        }
        else if(cal.get(Calendar.MONTH) == 6){
            period = 1;
        }
        else if(cal.get(Calendar.MONTH) == 7){
            period = 1;
        }
        else if(cal.get(Calendar.MONTH) == 8){
            period = 1;
        }
        //return period
        return period;
        
    }//end getDate()
    
    //checks if date cell is null or not
    public static Date nullDate(Cell c, String d) throws ParseException{

        Date date = new Date();
       //date format is <Day>.<Month>.<Year>
       SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
                if(c == null){
                    date = null;
                }
                else{
                    date = dateFormat.parse(c.toString());
                }
        return date;
    }
    
    //works the same way as getDate() in GetXlsx, WriteXlsx, GenerateXlsx except it takes a certain date as a parameter
    public static int periodChecker(Date datetime) throws ParseException{
        int period = 0;
        Calendar cal = Calendar.getInstance();
        cal.setTime(datetime);
        //FIRST SEMESTER CUT OFF DATES
        //gets month from given date
        //months start at 0, 0 being Jan
        //if month is October, November, December or Jan then period is 2
        if(cal.get(Calendar.MONTH) >= 9){
            period = 1;
        }
        else if(cal.get(Calendar.MONTH) == 0){
            period = 1;
        }
        
        
        //SECOND SEMESTER CUT OFF DATES
        //if month is Feb, March, April or May then period is 3
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
        //if month is June, July, Aug, September then period is 1
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
    }
    
  
}
