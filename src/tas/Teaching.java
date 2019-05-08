package tas;

import com.mongodb.DB;
import com.mongodb.DBCollection;
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
public class Teaching {
    
    
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
    
    public static void teaching() throws IOException, ParseException{
        //connect to local mongodb
        Mongo mongo = new Mongo("localhost", 27017);
        DB tas = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = tas.getCollection("TAS2");
        
        //remove all entries from the database
//        DBCursor rem = collection.find();
//        while(rem.hasNext()){
//            collection.remove(rem.next());
//        }
        //user chooses directory containing all users tas excel sheets
        
         //read excel file
                XSSFWorkbook wb = new XSSFWorkbook("TAS_WE.xlsx");
                
                for(int i = 4; i < 22; i++){
                    
                        
                        Cell cSur = wb.getSheetAt(0).getRow(i).getCell(0, xc.CREATE_NULL_AS_BLANK);
                        String sur = cSur.getStringCellValue();
                        System.out.println(sur);
                        
                        Cell cFir = wb.getSheetAt(0).getRow(i).getCell(1, xc.CREATE_NULL_AS_BLANK);
                        String fir = cFir.getStringCellValue();
                        System.out.println(fir);
                        
                        Cell cMc1 = wb.getSheetAt(0).getRow(i).getCell(2, xc.CREATE_NULL_AS_BLANK);
                        Double mc1 = cMc1.getNumericCellValue();
                        System.out.println(mc1);
                        
                        Cell cAss1 = wb.getSheetAt(0).getRow(i).getCell(3, xc.CREATE_NULL_AS_BLANK);
                        Double ass1 = cAss1.getNumericCellValue();
                        System.out.println(ass1);
                        
                        Cell cMc2 = wb.getSheetAt(0).getRow(i).getCell(4, xc.CREATE_NULL_AS_BLANK);
                        Double mc2 = cMc2.getNumericCellValue();
                        System.out.println(mc2);
                        
                        Cell cAss2 = wb.getSheetAt(0).getRow(i).getCell(5, xc.CREATE_NULL_AS_BLANK);
                        Double ass2 = cAss2.getNumericCellValue();
                        System.out.println(ass2);
                        
                        Cell cMc3 = wb.getSheetAt(0).getRow(i).getCell(6, xc.CREATE_NULL_AS_BLANK);
                        Double mc3 = cMc3.getNumericCellValue();
                        System.out.println(mc3);
                        
                        
                        Cell cGASept = wb.getSheetAt(0).getRow(i).getCell(7, xc.CREATE_NULL_AS_BLANK);
                        Double GASept = cGASept.getNumericCellValue();
                        System.out.println(GASept);
                        
                        
                        Cell cGADec = wb.getSheetAt(0).getRow(i).getCell(8, xc.CREATE_NULL_AS_BLANK);
                        Double GADec = cGADec.getNumericCellValue();
                        System.out.println(GADec);
                        
                        
                        Cell cGAMar = wb.getSheetAt(0).getRow(i).getCell(9, xc.CREATE_NULL_AS_BLANK);
                        Double GAMar = cGAMar.getNumericCellValue();
                        System.out.println(GAMar);
                        
                        
                        Cell cGAyr = wb.getSheetAt(0).getRow(i).getCell(10, xc.CREATE_NULL_AS_BLANK);
                        Double GAyr = cGAyr.getNumericCellValue();
                        System.out.println(GAyr);
                        
                        Cell cR = wb.getSheetAt(0).getRow(i).getCell(11, xc.CREATE_NULL_AS_BLANK);
                        Double r = cR.getNumericCellValue();
                        System.out.println(r);
                        
                        
                        Cell cCother = wb.getSheetAt(0).getRow(i).getCell(12, xc.CREATE_NULL_AS_BLANK);
                        Double cother = cCother.getNumericCellValue();
                        System.out.println(cother);
                        
                        
                        Cell cCotherSupp = wb.getSheetAt(0).getRow(i).getCell(13, xc.CREATE_NULL_AS_BLANK);
                        Double cOtherSupp = cCotherSupp.getNumericCellValue();
                        System.out.println(cOtherSupp);
                        
                        
                        Cell cOtherT = wb.getSheetAt(0).getRow(i).getCell(14, xc.CREATE_NULL_AS_BLANK);
                        Double otherT = cOtherT.getNumericCellValue();
                        System.out.println(otherT);
                        
                        
                        Cell cOtherpgr = wb.getSheetAt(0).getRow(i).getCell(15, xc.CREATE_NULL_AS_BLANK);
                        Double otherpgr = cOtherpgr.getNumericCellValue();
                        System.out.println(otherpgr);
                        
                        
                        Cell cOthersupp = wb.getSheetAt(0).getRow(i).getCell(16, xc.CREATE_NULL_AS_BLANK);
                        Double othersupp = cOthersupp.getNumericCellValue();
                        System.out.println(othersupp);
                        
                        
                        Cell cOtherfurther = wb.getSheetAt(0).getRow(i).getCell(17, xc.CREATE_NULL_AS_BLANK);
                        Double otherfurther = cOtherfurther.getNumericCellValue();
                        System.out.println(otherfurther);
                        
                        
                        Cell cOthermgmt = wb.getSheetAt(0).getRow(i).getCell(18, xc.CREATE_NULL_AS_BLANK);
                        Double othermgmt = cOthermgmt.getNumericCellValue();
                        System.out.println(othermgmt);
                        
                        
                        Cell ctime = wb.getSheetAt(0).getRow(i).getCell(19, xc.CREATE_NULL_AS_BLANK);
                        Double time = ctime.getNumericCellValue();
                        System.out.println(time);
                        
                        Cell cremain = wb.getSheetAt(0).getRow(i).getCell(20, xc.CREATE_NULL_AS_BLANK);
                        Double remain = cremain.getNumericCellValue();
                        System.out.println(remain);
                        
                        
                        System.out.println("------------------------------------------------");
                        System.out.println(sur + "\n" + fir + "\n" + mc1 + "\n" + ass1 + "\n" + mc2 + "\n" + ass2 + "\n" + mc3 + "\n" + GASept + "\n" + GADec + "\n" + GAMar + "\n" + GAyr + "\n" + r + "\n" + cother + "\n" + cOtherSupp + "\n" + otherT + "\n" + otherpgr + "\n" + othersupp + "\n" + otherfurther + "\n" + othermgmt + "\n" + time +"\n" + remain);      
                               
                        
                    
                }
                
    }
                
}
