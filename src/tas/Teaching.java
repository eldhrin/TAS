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
public class Teaching {
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
    
    public static void teaching() throws IOException, ParseException{
        //connect to local mongodb
        Mongo mongo = new Mongo("localhost", 27017);
        DB tas = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = tas.getCollection("TAS");
        
        //remove all entries from the database
        DBCursor rem = collection.find();
        while(rem.hasNext()){
            collection.remove(rem.next());
        }
        //user chooses directory containing all users tas excel sheets
        
         //read excel file
                XSSFWorkbook wb = new XSSFWorkbook("TAS_WE.xlsx");                             
                
                for(int i = 4; i < 32; i++){
                    
                    BasicDBObject document = new BasicDBObject();
                        
                        Cell cSur = wb.getSheetAt(0).getRow(i).getCell(0, xc.CREATE_NULL_AS_BLANK);
                        String sur = cSur.getStringCellValue();
                        System.out.println(sur);
                        
                        
                        Cell cFir = wb.getSheetAt(0).getRow(i).getCell(1, xc.CREATE_NULL_AS_BLANK);
                        String fir = cFir.getStringCellValue();
                        System.out.println(fir);
                        document.put("name", fir + " " + sur);
                        
                        Cell cID = wb.getSheetAt(0).getRow(i).getCell(2, xc.CREATE_NULL_AS_BLANK);
                        String id = cID.getStringCellValue();
                        id = id.toUpperCase();
                        System.out.println(id);
                        document.put("id", id);
                        
                        Cell cMc1 = wb.getSheetAt(0).getRow(i).getCell(3, xc.CREATE_NULL_AS_BLANK);
                        Double mc1 = cMc1.getNumericCellValue();
                        Double mc1L = 0.0;
                        mc1 = mc1*p;
                        mc1L = mc1;
                        mc1 = mc1*90;
                        mc1 = mc1/10000;
                        mc1L = mc1L * 10;
                        mc1L = mc1L/10000;
                        System.out.println(mc1);
                        System.out.println(mc1L);
                        document.put("module co-ord", mc1);
                        document.put("module supp", mc1L);
                        
                        Cell cAss1 = wb.getSheetAt(0).getRow(i).getCell(4, xc.CREATE_NULL_AS_BLANK);
                        Double ass1 = cAss1.getNumericCellValue();
                        ass1 = ass1*p;
                        ass1 = ass1*90;
                        ass1 = ass1/10000;
                        System.out.println(ass1);
                        document.put("assist", ass1);
                        
                        Cell cMc2 = wb.getSheetAt(0).getRow(i).getCell(5, xc.CREATE_NULL_AS_BLANK);
                        Double mc2 = cMc2.getNumericCellValue();
                        mc2 = mc2*p;
                        Double mc2L = mc2;
                        mc2 = mc2*90;
                        mc2 = mc2/10000;
                        mc2L = mc2L*10;
                        mc2L = mc2L/10000; 
                        System.out.println(mc2);
                        System.out.println(mc2L);
                        document.put("module co-ord sem2", mc2);
                        document.put("module supp sem2", mc2L);
                        
                        Cell cAss2 = wb.getSheetAt(0).getRow(i).getCell(6, xc.CREATE_NULL_AS_BLANK);
                        Double ass2 = cAss2.getNumericCellValue();
                        ass2 = ass2*p;
                        ass2 = ass2*90;
                        ass2 = ass2/10000;
                        System.out.println(ass2);
                        document.put("massist sem2", ass2);
                        
                        Cell cMc3 = wb.getSheetAt(0).getRow(i).getCell(7, xc.CREATE_NULL_AS_BLANK);
                        Double mc3 = cMc3.getNumericCellValue();
                        mc3 = mc3*p;
                        Double mc3L = mc3;
                        mc3 = mc3*90;
                        mc3 = mc3/10000;
                        mc3L = mc3L*10;
                        mc3L = mc3L/10000;
                        System.out.println(mc3);
                        System.out.println(mc3L);
                        document.put("module co-ord sem3", mc3);
                        document.put("module supp sem3", mc3L);
                        
                        
                        Cell cGASept = wb.getSheetAt(0).getRow(i).getCell(8, xc.CREATE_NULL_AS_BLANK);
                        Double GASept = cGASept.getNumericCellValue();
                        GASept = GASept*p;
                        Double GASeptL = GASept;
                        GASept = GASept*90;
                        GASept = GASept/10000;
                        GASeptL = GASeptL*10;
                        GASeptL = GASeptL/10000;
                        System.out.println(GASept);
                        System.out.println(GASeptL);
                        document.put("GA SEPT", GASept);
                        document.put("GA SEPT L", GASeptL);
                        
                        
                        Cell cGADec = wb.getSheetAt(0).getRow(i).getCell(9, xc.CREATE_NULL_AS_BLANK);
                        Double GADec = cGADec.getNumericCellValue();
                        GADec = GADec*p;
                        Double GADecL = GADec;
                        GADec = GADec*90;
                        GADec = GADec/10000;
                        GADecL = GADecL*10;
                        GADecL = GADecL/10000;
                        System.out.println(GADec);
                        System.out.println(GADecL);
                        document.put("GA DEC", GADec);
                        document.put("GA DEC L", GADecL);
                        
                        
                        Cell cGAMar = wb.getSheetAt(0).getRow(i).getCell(10, xc.CREATE_NULL_AS_BLANK);
                        Double GAMar = cGAMar.getNumericCellValue();
                        GAMar = GAMar*p;
                        Double GAMarL = GAMar;
                        GAMar = GAMar*90;
                        GAMar = GAMar/10000;
                        GAMarL = GAMarL*10;
                        GAMarL = GAMarL/10000;
                        System.out.println(GAMar);
                        System.out.println(GAMarL);
                        document.put("GA MAR", GAMar);
                        document.put("GA MAR L", GAMarL);
                        
                        
                        Cell cGAyr = wb.getSheetAt(0).getRow(i).getCell(11, xc.CREATE_NULL_AS_BLANK);
                        Double GAyr = cGAyr.getNumericCellValue();
                        GAyr = GAyr*p;
                        Double GAyrL = GAyr;
                        GAyr = GAyr*10;
                        GAyr = GAyr/10000;
                        GAyrL = GAyrL*10;
                        GAyrL = GAyrL/10000;
                        System.out.println(GAyr);
                        System.out.println(GAyrL);
                        document.put("GA YEAR", GAyr);
                        document.put("GA YEAR L", GAyrL);
                        
                        Cell cR = wb.getSheetAt(0).getRow(i).getCell(12, xc.CREATE_NULL_AS_BLANK);
                        Double r = cR.getNumericCellValue();
                        r = r*p;
                        Double rL = r;
                        r = r*90;
                        r = r/10000;
                        rL = rL*10;
                        rL = rL/10000;
                        System.out.println(r);
                        System.out.println(rL);
                        document.put("research", r);
                        document.put("research L", rL);
                        
                        
                        Cell cCother = wb.getSheetAt(0).getRow(i).getCell(13, xc.CREATE_NULL_AS_BLANK);
                        Double cother = cCother.getNumericCellValue();
                        cother = cother*p;
                        Double cotherL = cother;
                        cother = cother*90;
                        cother = cother/10000;
                        cotherL = cotherL*10;
                        cotherL = cotherL/10000;
                        System.out.println(cother);
                        System.out.println(cotherL);
                        document.put("c other", cother);
                        document.put("cOther L", cotherL);
                        
                        
                        Cell cCotherSupp = wb.getSheetAt(0).getRow(i).getCell(14, xc.CREATE_NULL_AS_BLANK);
                        Double cOtherSupp = cCotherSupp.getNumericCellValue();
                        cOtherSupp = cOtherSupp*p;
                        Double cOtherSuppL = cOtherSupp;
                        cOtherSupp = cOtherSupp*80;
                        cOtherSupp = cOtherSupp/10000;
                        cOtherSuppL = cOtherSuppL*20;
                        cOtherSuppL = cOtherSuppL/10000;
                        System.out.println(cOtherSupp);
                        System.out.println(cOtherSuppL);
                        document.put("cother supp", cOtherSupp);
                        document.put("cother supp L", cOtherSuppL);
                        
                        
                        Cell cOtherT = wb.getSheetAt(0).getRow(i).getCell(15, xc.CREATE_NULL_AS_BLANK);
                        Double otherT = cOtherT.getNumericCellValue();
                        otherT = otherT*p/100;
                        System.out.println(otherT);
                        document.put("otherT", otherT);
                        
                        
                        Cell cOtherpgr = wb.getSheetAt(0).getRow(i).getCell(16, xc.CREATE_NULL_AS_BLANK);
                        Double otherpgr = cOtherpgr.getNumericCellValue();
                        otherpgr = otherpgr*p/100;
                        System.out.println(otherpgr);
                        document.put("pgr", otherpgr);
                        
                        
                        Cell cOthersupp = wb.getSheetAt(0).getRow(i).getCell(17, xc.CREATE_NULL_AS_BLANK);
                        Double othersupp = cOthersupp.getNumericCellValue();
                        othersupp = othersupp*p/100;
                        System.out.println(othersupp);
                        document.put("other supp", othersupp);
                        
                        
                        Cell cOtherfurther = wb.getSheetAt(0).getRow(i).getCell(18, xc.CREATE_NULL_AS_BLANK);
                        Double otherfurther = cOtherfurther.getNumericCellValue();
                        otherfurther = otherfurther*p/100;
                        System.out.println(otherfurther);
                        document.put("further", otherfurther);
                        
                        Cell cOthermgmt = wb.getSheetAt(0).getRow(i).getCell(19, xc.CREATE_NULL_AS_BLANK);
                        Double othermgmt = cOthermgmt.getNumericCellValue();
                        othermgmt = othermgmt*p/100;
                        System.out.println(othermgmt);
                        document.put("mgmt", othermgmt);
                        
                        
                        Cell ctime = wb.getSheetAt(0).getRow(i).getCell(20, xc.CREATE_NULL_AS_BLANK);
                        Double time = ctime.getNumericCellValue();
                        time = time*5;
                        time = time/100;
                        System.out.println(time);
                        document.put("time", time);
                        
                        Cell cremain = wb.getSheetAt(0).getRow(i).getCell(21, xc.CREATE_NULL_AS_BLANK);
                        Double remain = cremain.getNumericCellValue();
                        System.out.println(remain);
                        document.put("remain", remain);
                        
                        document.put("sem1", 0.0);
                        document.put("sem2", 0.0);
                        document.put("sem3", 0.0);
                        
                        collection.insert(document);
                        
                        
                        System.out.println("------------------------------------------------");
                        System.out.println(sur + "\n" + fir + "\n" + mc1 + "\n" + ass1 + "\n" + mc2 + "\n" + ass2 + "\n" + mc3 + "\n" + GASept + "\n" + GADec + "\n" + GAMar + "\n" + GAyr + "\n" + r + "\n" + cother + "\n" + cOtherSupp + "\n" + otherT + "\n" + otherpgr + "\n" + othersupp + "\n" + otherfurther + "\n" + othermgmt + "\n" + time +"\n" + remain);      
                               
                    
                }
                
    }
                
}
