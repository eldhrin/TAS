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
import com.mongodb.DBObject;
import com.mongodb.Mongo;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author fl8328
 */
public class WriteReportResearch {
    
    private static Row.MissingCellPolicy xc;
    
    public static XSSFWorkbook writeReport() throws IOException{
        
        //connect to the database
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //count is first row where we write the data to
        //find collection TAS
        DBCollection workload = db.getCollection("TAS_WL");
        DBCollection collection = db.getCollection("TAS_PGR");
        DBCollection research = db.getCollection("RESEARCH");
        DBCursor cursor = workload.find(); 
        DBCursor cur = collection.find();
        DBCursor res = research.find();
        //get number of entries in the database (this starts at 1)
        int dbcount = (int)workload.count();
        int dbc = (int)collection.count();
        int rescount = (int)research.count();
        //base file the program writes to
        FileOutputStream fileOut = new FileOutputStream("megareportKEEP.xlsx");
        //file that the program outputs (creates this file if it does not exist)
        String pathName = "H:\\NetBeansProjects\\TAS\\TAS-VD.xlsx";
        XSSFWorkbook wb = new XSSFWorkbook(pathName);
        int count = 6;
        int j = 0;
        int x = 0;
        int c = 0;
        double pt5 = 1/.5;
        double pt8 = 1/.8;

//        for(int x = 0; x < rescount; x++){
//            DBObject r = res.next();
            
                


                    for(j = 0; j < rescount; j++){
                      
                        
                    
                    DBObject r = res.next();
                    DBObject rc = (DBObject) r.get("cate");
                        System.out.println(j);
                        System.out.println("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx");
                        for(x = count; x < 34; x++){
                            Cell cLName = wb.getSheetAt(1).getRow(x).getCell(0, xc.CREATE_NULL_AS_BLANK);
                            String lName = new String();
                            System.out.println(x);
                            lName = Null.nullString(cLName, lName);
                            
                            if(r.get("projLead").equals(lName.toUpperCase())){
                                System.out.println(lName);
                                    String cate = rc.get("0").toString();
                                    if(cate.contains("Research Councils") || cate.contains("research councils") || cate.contains("RESEARCH COUNCILS")){
                                        Cell cRC = wb.getSheetAt(1).getRow(x).getCell(5, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cRC.getNumericCellValue();
                                        Double val = cRC.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("1").toString());
                                        }
                                        cRC.setCellValue(val);
                                        System.out.println("HERE"); 
                                        val = 0.0;
                                    }
                                    
                                    
                                    else if(cate.contains("UK Govt Depts") || cate.contains("uk govt depts") || cate.contains("UK GOVT DEPTS") || cate.contains("UK Government Departments") || cate.contains("uk government departments") || cate.contains("UK GOVERNMENT DEPARTMENTS")){
                                        Cell cUKD = wb.getSheetAt(1).getRow(x).getCell(6, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKD.getNumericCellValue();
                                        Double val = cUKD.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("1").toString());
                                        }
                                        cUKD.setCellValue(val);
                                        System.out.println("HERE");
                                        val = 0.0;
                                    }
                                    
                                    else if(cate.contains("European Commission") || cate.contains("european commission") || cate.contains("EUROPEAN COMMISSION") || cate.contains("EU Commission") || cate.contains("eu commission") || cate.contains("EU COMMISSION")){
                                        Cell cEU = wb.getSheetAt(1).getRow(x).getCell(7, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cEU.getNumericCellValue();
                                        Double val = cEU.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("1").toString());
                                        }
                                        cEU.setCellValue(val);
                                        System.out.println("HERE");   
                                        val = 0.0;
                                    }
                                    
                                    else if(cate.contains("UK Charities") || cate.contains("uk charities") || cate.contains("UK CHARITIES")){
                                        Cell cUKC = wb.getSheetAt(1).getRow(x).getCell(8, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKC.getNumericCellValue();
                                        Double val = cUKC.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("1").toString());
                                        }
                                        cUKC.setCellValue(val);
                                        System.out.println("HERE");  
                                        val = 0.0;
                                    }
                                    
                                    else if(cate.contains("UK Industry") || cate.contains("uk industry") || cate.contains("UK INDUSTRY")){
                                        Cell cUKI = wb.getSheetAt(1).getRow(x).getCell(9, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKI.getNumericCellValue();
                                        Double val = cUKI.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("1").toString());
                                        }
                                        cUKI.setCellValue(val);
                                        System.out.println("HERE");  
                                        val = 0.0;
                                    }
                                    
                                    else if(cate.contains("KTPs") || cate.contains("ktps") || cate.contains("KTPS") || cate.contains("KTP") || cate.contains("ktp")){
                                        Cell cKTP = wb.getSheetAt(1).getRow(x).getCell(10, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cKTP.getNumericCellValue();
                                        Double val = cKTP.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("1").toString());
                                        }
                                        cKTP.setCellValue(val);
                                        System.out.println("HERE");  
                                        val = 0.0;
                                    }
                                    
                                    else if(cate.contains("Other") || cate.contains("OTHER") || cate.contains("other")){
                                        Cell cOther = wb.getSheetAt(1).getRow(x).getCell(11, xc.CREATE_NULL_AS_BLANK);
                                       Double rcount = 0.0;
                                        rcount = cOther.getNumericCellValue();
                                        Double val = cOther.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("1").toString());
                                        }
                                        cOther.setCellValue(val);
                                        System.out.println("HERE");  
                                        val = 0.0;
                                    }
                                    
                                    else if(cate.contains("SFC") || cate.contains("sfc") || cate.contains("Sfc")){
                                        Cell cSFC = wb.getSheetAt(1).getRow(x).getCell(13, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cSFC.getNumericCellValue();
                                        Double val = cSFC.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("1").toString());
                                        }
                                        cSFC.setCellValue(val);
                                        System.out.println("HERE"); 
                                        val = 0.0;
                                    }
                                    
                                    else if(cate.contains("INTERNALLY") || cate.contains("Internally") || cate.contains("internally")){
                                        Cell cInt = wb.getSheetAt(1).getRow(x).getCell(15, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cInt.getNumericCellValue();
                                        Double val = cInt.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("1").toString());
                                        }
                                        cInt.setCellValue(val);
                                        System.out.println("HERE");
                                        val = 0.0;

                                    }
                                
                            }
                            
                            
                            
                            
                            
                            else if(r.get("co inv 1").equals(lName.toUpperCase())){
                                System.out.println(lName);
                                    String cate1 = rc.get("0").toString();
                                    if(cate1.contains("Research Councils") || cate1.contains("research councils") || cate1.contains("RESEARCH COUNCILS")){
                                        Cell cRC = wb.getSheetAt(1).getRow(x).getCell(5, xc.CREATE_NULL_AS_BLANK);
                                       Double rcount = 0.0;
                                        rcount = cRC.getNumericCellValue();
                                        Double val = cRC.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte2").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("2").toString());
                                        }
                                        cRC.setCellValue(val);
                                        System.out.println("HERE"); 
                                         val = 0.0;
                                    }
                                    
                                    
                                    else if(cate1.contains("UK Govt Depts") || cate1.contains("uk govt depts") || cate1.contains("UK GOVT DEPTS") || cate1.contains("UK Government Departments") || cate1.contains("uk government departments") || cate1.contains("UK GOVERNMENT DEPARTMENTS")){
                                        Cell cUKD = wb.getSheetAt(1).getRow(x).getCell(6, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKD.getNumericCellValue();
                                        Double val = cUKD.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte2").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("2").toString());
                                        }
                                        cUKD.setCellValue(val);
                                        System.out.println("HERE");  
                                        val = 0.0;
                                    }
                                    
                                    else if(cate1.contains("European Commission") || cate1.contains("european commission") || cate1.contains("EUROPEAN COMMISSION") || cate1.contains("EU Commission") || cate1.contains("eu commission") || cate1.contains("EU COMMISSION")){
                                        Cell cEU = wb.getSheetAt(1).getRow(x).getCell(7, xc.CREATE_NULL_AS_BLANK); 
                                        Double rcount = 0.0;
                                        rcount = cEU.getNumericCellValue();
                                        Double val = cEU.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte2").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("2").toString());
                                        }
                                        cEU.setCellValue(val);
                                        System.out.println("HERE");   
                                        val = 0.0;
                                    }
                                    
                                    else if(cate1.contains("UK Charities") || cate1.contains("uk charities") || cate1.contains("UK CHARITIES")){
                                        Cell cUKC = wb.getSheetAt(1).getRow(x).getCell(8, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKC.getNumericCellValue();
                                        Double val = cUKC.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte2").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("2").toString());
                                        }
                                        cUKC.setCellValue(val);
                                        System.out.println("HERE");    
                                        val = 0.0;
                                    }
                                    
                                    else if(cate1.contains("UK Industry") || cate1.contains("uk industry") || cate1.contains("UK INDUSTRY")){
                                        Cell cUKI = wb.getSheetAt(1).getRow(x).getCell(9, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKI.getNumericCellValue();
                                                Double val = cUKI.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte2").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("2").toString());
                                        }
                                        cUKI.setCellValue(val);
                                        System.out.println("HERE");  
                                        val = 0.0;
                                    }
                                    
                                    else if(cate1.contains("KTPs") || cate1.contains("ktps") || cate1.contains("KTPS") || cate1.contains("KTP") || cate1.contains("ktp")){
                                        Cell cKTP = wb.getSheetAt(1).getRow(x).getCell(10, xc.CREATE_NULL_AS_BLANK); 
                                         Double rcount = 0.0;
                                        rcount = cKTP.getNumericCellValue();
                                        Double val = cKTP.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte2").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("2").toString());
                                        }
                                        cKTP.setCellValue(val);
                                        System.out.println("HERE");
                                        val = 0.0;
                                    }
                                    
                                    else if(cate1.contains("Other") || cate1.contains("OTHER") || cate1.contains("other")){
                                        Cell cOther = wb.getSheetAt(1).getRow(x).getCell(11, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cOther.getNumericCellValue();
                                        Double val = cOther.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte2").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("2").toString());
                                        }
                                        cOther.setCellValue(val);
                                        System.out.println("HERE");
                                        val = 0.0;
                                    }
                                    
                                    else if(cate1.contains("SFC") || cate1.contains("sfc") || cate1.contains("Sfc")){
                                        Cell cSFC = wb.getSheetAt(1).getRow(x).getCell(13, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cSFC.getNumericCellValue();
                                        Double val = cSFC.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte2").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("2").toString());
                                        }
                                                cSFC.setCellValue(val);
                                                System.out.println("HERE");
                                                val = 0.0;
                                    }
                                    
                                    else if(cate1.contains("INTERNALLY") || cate1.contains("Internally") || cate1.contains("internally")){
                                        Cell cInt = wb.getSheetAt(1).getRow(x).getCell(15, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cInt.getNumericCellValue();
                                        Double val = cInt.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte2").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("2").toString());
                                        }
                                        cInt.setCellValue(val);
                                        System.out.println("HERE"); 
                                        val = 0.0;
                                    }
                                
                            }
                            

                            
                            else if(lName.toUpperCase().contains(r.get("co inv 2").toString())){
                                System.out.println(lName);
                                    String cate2 = rc.get("0").toString();
                                    if(cate2.contains("Research Councils") || cate2.contains("research councils") || cate2.contains("RESEARCH COUNCILS")){
                                        Cell cRC = wb.getSheetAt(1).getRow(x).getCell(5, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cRC.getNumericCellValue();
                                        Double val = cRC.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte3").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("3").toString());
                                        }
                                        cRC.setCellValue(val);
                                        System.out.println("HERE"); 
                                        val = 0.0;
                                    }
                                    
                                    else if(cate2.contains("UK Govt Depts") || cate2.contains("uk govt depts") || cate2.contains("UK GOVT DEPTS") || cate2.contains("UK Government Departments") || cate2.contains("uk government departments") || cate2.contains("UK GOVERNMENT DEPARTMENTS")){
                                        Cell cUKD = wb.getSheetAt(1).getRow(x).getCell(6, xc.CREATE_NULL_AS_BLANK); 
                                        Double rcount = 0.0;
                                        rcount = cUKD.getNumericCellValue();
                                        Double val = cUKD.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte3").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("3").toString());
                                        }
                                        cUKD.setCellValue(val);
                                        System.out.println("HERE");   
                                        val = 0.0;
                                    }
                                    
                                    else if(cate2.contains("European Commission") || cate2.contains("european commission") || cate2.contains("EUROPEAN COMMISSION") || cate2.contains("EU Commission") || cate2.contains("eu commission") || cate2.contains("EU COMMISSION")){
                                        Cell cEU = wb.getSheetAt(1).getRow(x).getCell(7, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cEU.getNumericCellValue();
                                        Double val = cEU.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte3").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("3").toString());
                                        }
                                        cEU.setCellValue(val);
                                        System.out.println("HERE");    
                                        val = 0.0;
                                    }
                                    
                                    else if(cate2.contains("UK Charities") || cate2.contains("uk charities") || cate2.contains("UK CHARITIES")){
                                        Cell cUKC = wb.getSheetAt(1).getRow(x).getCell(8, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKC.getNumericCellValue();
                                        Double val = cUKC.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte3").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("3").toString());
                                        }
                                        cUKC.setCellValue(val);
                                        System.out.println("HERE"); 
                                        val = 0.0;
                                    }
                                    
                                    else if(cate2.contains("UK Industry") || cate2.contains("uk industry") || cate2.contains("UK INDUSTRY")){
                                        Cell cUKI = wb.getSheetAt(1).getRow(x).getCell(9, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKI.getNumericCellValue();
                                        Double val = cUKI.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte3").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("3").toString());
                                        }
                                        cUKI.setCellValue(val);
                                        System.out.println("HERE");  
                                        val = 0.0;
                                    }
                                    
                                    else if(cate2.contains("KTPs") || cate2.contains("ktps") || cate2.contains("KTPS") || cate2.contains("KTP") || cate2.contains("ktp")){
                                        Cell cKTP = wb.getSheetAt(1).getRow(x).getCell(10, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cKTP.getNumericCellValue();
                                        Double val = cKTP.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte3").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("3").toString());
                                        }
                                        cKTP.setCellValue(val);
                                        System.out.println("HERE"); 
                                        val = 0.0;
                                    }
                                    
                                    else if(cate2.contains("Other") || cate2.contains("OTHER") || cate2.contains("other")){
                                        Cell cOther = wb.getSheetAt(1).getRow(x).getCell(11, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cOther.getNumericCellValue();
                                        Double val = cOther.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte3").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("3").toString());
                                        }
                                        cOther.setCellValue(val);
                                        System.out.println("HERE");
                                        val = 0.0;
                                    }
                                    
                                    else if(cate2.contains("SFC") || cate2.contains("sfc") || cate2.contains("Sfc")){
                                        Cell cSFC = wb.getSheetAt(1).getRow(x).getCell(13, xc.CREATE_NULL_AS_BLANK); 
                                        Double rcount = 0.0;
                                        rcount = cSFC.getNumericCellValue();
                                        Double val = cSFC.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte3").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("3").toString());
                                        }
                                        cSFC.setCellValue(val);
                                        System.out.println("HERE");
                                        val = 0.0;
                                    }
                                    
                                    else if(cate2.contains("INTERNALLY") || cate2.contains("Internally") || cate2.contains("internally")){
                                        Cell cInt = wb.getSheetAt(1).getRow(x).getCell(15, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cInt.getNumericCellValue();
                                        Double val = cInt.getNumericCellValue();
                                        if(rc.get("4").equals("")){
                                            val = val + Double.parseDouble(r.get("fte3").toString());
                                        }
                                        else{
                                            val = val + Double.parseDouble(rc.get("3").toString());
                                        }
                                        cInt.setCellValue(val);
                                        System.out.println("HERE"); 
                                        val = 0.0;
                                                

                                    }
                                
                                //-----------------------------------------------------------------------------------
                            
                            }
                            
                            if(rc.get("4") != ""){
                            
                                if(r.get("projLead").equals(lName.toUpperCase())){
                                System.out.println(lName);
                                    String cate = rc.get("4").toString();
                                    if(cate.contains("Research Councils") || cate.contains("research councils") || cate.contains("RESEARCH COUNCILS")){
                                        Cell cRC = wb.getSheetAt(1).getRow(x).getCell(5, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cRC.getNumericCellValue();
                                        Double val = cRC.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("5").toString());
                                        cRC.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate.contains("UK Govt Depts") || cate.contains("uk govt depts") || cate.contains("UK GOVT DEPTS") || cate.contains("UK Government Departments") || cate.contains("uk government departments") || cate.contains("UK GOVERNMENT DEPARTMENTS")){
                                        Cell cUKD = wb.getSheetAt(1).getRow(x).getCell(6, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKD.getNumericCellValue();
                                        Double val = cUKD.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("5").toString());
                                        cUKD.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    else if(cate.contains("European Commission") || cate.contains("european commission") || cate.contains("EUROPEAN COMMISSION") || cate.contains("EU Commission") || cate.contains("eu commission") || cate.contains("EU COMMISSION")){
                                        Cell cEU = wb.getSheetAt(1).getRow(x).getCell(7, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cEU.getNumericCellValue();
                                        Double val = cEU.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("5").toString());
                                        cEU.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate.contains("UK Charities") || cate.contains("uk charities") || cate.contains("UK CHARITIES")){
                                        Cell cUKC = wb.getSheetAt(1).getRow(x).getCell(8, xc.CREATE_NULL_AS_BLANK); 
                                        Double rcount = 0.0;
                                        rcount = cUKC.getNumericCellValue();
                                        Double val = cUKC.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("5").toString());
                                        cUKC.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate.contains("UK Industry") || cate.contains("uk industry") || cate.contains("UK INDUSTRY")){
                                        Cell cUKI = wb.getSheetAt(1).getRow(x).getCell(9, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKI.getNumericCellValue();
                                        Double val = cUKI.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("5").toString());
                                        cUKI.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate.contains("KTPs") || cate.contains("ktps") || cate.contains("KTPS") || cate.contains("KTP") || cate.contains("ktp")){
                                        Cell cKTP = wb.getSheetAt(1).getRow(x).getCell(10, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cKTP.getNumericCellValue();
                                        Double val = cKTP.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("5").toString());
                                        cKTP.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate.contains("Other") || cate.contains("OTHER") || cate.contains("other")){
                                        Cell cOther = wb.getSheetAt(1).getRow(x).getCell(11, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cOther.getNumericCellValue();
                                        Double val = cOther.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("5").toString());
                                        cOther.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate.contains("SFC") || cate.contains("sfc") || cate.contains("Sfc")){
                                        Cell cSFC = wb.getSheetAt(1).getRow(x).getCell(13, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cSFC.getNumericCellValue();
                                        Double val = cSFC.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("5").toString());
                                       cSFC.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate.contains("INTERNALLY") || cate.contains("Internally") || cate.contains("internally")){
                                        Cell cInt = wb.getSheetAt(1).getRow(x).getCell(15, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cInt.getNumericCellValue();
                                        Double val = cInt.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("5").toString());
                                        cInt.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                
                            }  
                            
                            else if(r.get("co inv 1").equals(lName.toUpperCase())){
                                System.out.println(lName);
                                    String cate1 = rc.get("4").toString();
                                    if(cate1.contains("Research Councils") || cate1.contains("research councils") || cate1.contains("RESEARCH COUNCILS")){
                                        Cell cRC = wb.getSheetAt(1).getRow(x).getCell(5, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cRC.getNumericCellValue();
                                        Double val = cRC.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("6").toString());
                                        cRC.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate1.contains("UK Govt Depts") || cate1.contains("uk govt depts") || cate1.contains("UK GOVT DEPTS") || cate1.contains("UK Government Departments") || cate1.contains("uk government departments") || cate1.contains("UK GOVERNMENT DEPARTMENTS")){
                                        Cell cUKD = wb.getSheetAt(1).getRow(x).getCell(6, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKD.getNumericCellValue();
                                        Double val = cUKD.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("6").toString());
                                        cUKD.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate1.contains("European Commission") || cate1.contains("european commission") || cate1.contains("EUROPEAN COMMISSION") || cate1.contains("EU Commission") || cate1.contains("eu commission") || cate1.contains("EU COMMISSION")){
                                        Cell cEU = wb.getSheetAt(1).getRow(x).getCell(7, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cEU.getNumericCellValue();
                                        Double val = cEU.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("6").toString());
                                        cEU.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate1.contains("UK Charities") || cate1.contains("uk charities") || cate1.contains("UK CHARITIES")){
                                        Cell cUKC = wb.getSheetAt(1).getRow(x).getCell(8, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKC.getNumericCellValue();
                                        Double val = cUKC.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("6").toString());
                                        cUKC.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate1.contains("UK Industry") || cate1.contains("uk industry") || cate1.contains("UK INDUSTRY")){
                                        Cell cUKI = wb.getSheetAt(1).getRow(x).getCell(9, xc.CREATE_NULL_AS_BLANK); 
                                        Double rcount = 0.0;
                                        rcount = cUKI.getNumericCellValue();
                                        Double val = cUKI.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("6").toString());
                                        cUKI.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate1.contains("KTPs") || cate1.contains("ktps") || cate1.contains("KTPS") || cate1.contains("KTP") || cate1.contains("ktp")){
                                        Cell cKTP = wb.getSheetAt(1).getRow(x).getCell(10, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cKTP.getNumericCellValue();
                                        Double val = cKTP.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("6").toString());
                                        cKTP.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate1.contains("Other") || cate1.contains("OTHER") || cate1.contains("other")){
                                        Cell cOther = wb.getSheetAt(1).getRow(x).getCell(11, xc.CREATE_NULL_AS_BLANK); 
                                        Double rcount = 0.0;
                                        rcount = cOther.getNumericCellValue();
                                        Double val = cOther.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("6").toString());
                                        cOther.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate1.contains("SFC") || cate1.contains("sfc") || cate1.contains("Sfc")){
                                        Cell cSFC = wb.getSheetAt(1).getRow(x).getCell(13, xc.CREATE_NULL_AS_BLANK); 
                                        Double rcount = 0.0;
                                        rcount = cSFC.getNumericCellValue();
                                        Double val = cSFC.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("6").toString());
                                        cSFC.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate1.contains("INTERNALLY") || cate1.contains("Internally") || cate1.contains("internally")){
                                        Cell cInt = wb.getSheetAt(1).getRow(x).getCell(15, xc.CREATE_NULL_AS_BLANK); 
                                        Double rcount = 0.0;
                                        rcount = cInt.getNumericCellValue();
                                        Double val = cInt.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("6").toString());
                                        cInt.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                
                            }
                            

                            
                            else if(lName.toUpperCase().contains(r.get("co inv 2").toString())){
                                System.out.println(lName);
                                    String cate2 = rc.get("4").toString();
                                    if(cate2.contains("Research Councils") || cate2.contains("research councils") || cate2.contains("RESEARCH COUNCILS")){
                                        Cell cRC = wb.getSheetAt(1).getRow(x).getCell(5, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cRC.getNumericCellValue();
                                        Double val = cRC.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("7").toString());
                                        cRC.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate2.contains("UK Govt Depts") || cate2.contains("uk govt depts") || cate2.contains("UK GOVT DEPTS") || cate2.contains("UK Government Departments") || cate2.contains("uk government departments") || cate2.contains("UK GOVERNMENT DEPARTMENTS")){
                                        Cell cUKD = wb.getSheetAt(1).getRow(x).getCell(6, xc.CREATE_NULL_AS_BLANK); 
                                        Double rcount = 0.0;
                                        rcount = cUKD.getNumericCellValue();
                                        Double val = cUKD.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("7").toString());
                                        cUKD.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate2.contains("European Commission") || cate2.contains("european commission") || cate2.contains("EUROPEAN COMMISSION") || cate2.contains("EU Commission") || cate2.contains("eu commission") || cate2.contains("EU COMMISSION")){
                                        Cell cEU = wb.getSheetAt(1).getRow(x).getCell(7, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cEU.getNumericCellValue();
                                        Double val = cEU.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("7").toString());
                                        cEU.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate2.contains("UK Charities") || cate2.contains("uk charities") || cate2.contains("UK CHARITIES")){
                                        Cell cUKC = wb.getSheetAt(1).getRow(x).getCell(8, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKC.getNumericCellValue();
                                        Double val = cUKC.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("7").toString());
                                        cUKC.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate2.contains("UK Industry") || cate2.contains("uk industry") || cate2.contains("UK INDUSTRY")){
                                        Cell cUKI = wb.getSheetAt(1).getRow(x).getCell(9, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cUKI.getNumericCellValue();
                                        Double val = cUKI.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("7").toString());
                                        cUKI.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate2.contains("KTPs") || cate2.contains("ktps") || cate2.contains("KTPS") || cate2.contains("KTP") || cate2.contains("ktp")){
                                        Cell cKTP = wb.getSheetAt(1).getRow(x).getCell(10, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cKTP.getNumericCellValue();
                                        Double val = cKTP.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("7").toString());
                                        cKTP.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate2.contains("Other") || cate2.contains("OTHER") || cate2.contains("other")){
                                        Cell cOther = wb.getSheetAt(1).getRow(x).getCell(11, xc.CREATE_NULL_AS_BLANK); 
                                        Double rcount = 0.0;
                                        rcount = cOther.getNumericCellValue();
                                        Double val = cOther.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("7").toString());
                                        cOther.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate2.contains("SFC") || cate2.contains("sfc") || cate2.contains("Sfc")){
                                        Cell cSFC = wb.getSheetAt(1).getRow(x).getCell(13, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cSFC.getNumericCellValue();
                                        Double val = cSFC.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("7").toString());
                                        cSFC.setCellValue(val);
                                        System.out.println("HERE");
                                    }
                                    
                                    else if(cate2.contains("INTERNALLY") || cate2.contains("Internally") || cate2.contains("internally")){
                                        Cell cInt = wb.getSheetAt(1).getRow(x).getCell(15, xc.CREATE_NULL_AS_BLANK);
                                        Double rcount = 0.0;
                                        rcount = cInt.getNumericCellValue();
                                        Double val = cInt.getNumericCellValue();
                                        val = val + Double.parseDouble(rc.get("7").toString());   
                                        cInt.setCellValue(val);
                                        System.out.println("HERE");

                                    }
                                
                                //-----------------------------------------------------------------------------------
                            
                              }
                            }
                            
                            

                        
                        System.out.println("--------------------------------------------------------");

                
                    }
                    
                
        
            }
          wb.write(fileOut); 
        
        
          fileOut.close();
          //close workbook
          wb.close();
          return null;
    }
}
            

        

            
            
    
    

