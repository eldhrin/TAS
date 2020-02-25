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
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author fl8328
 */
public class Research {
    
    public static float nullFloat(Cell c, float fl){
        if(c == null){
            fl = 0.00000f;
        }
        //if DoubleCell != blank, d = value of cell
        else{
            fl = Float.parseFloat(c.toString());
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
    
    public static void research() throws IOException, ParseException{
        //connect to local mongodb
        Mongo mongo = new Mongo("localhost", 27017);
        DB tas = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = tas.getCollection("RESEARCH");
        
        //remove all entries from the database
        DBCursor rem = collection.find();
        while(rem.hasNext()){
            collection.remove(rem.next());
        }
        int in = 0;
        String addChar = "";
        Double charper = 0.0;
        Double percent = 0.0;
        //user chooses directory containing all users tas excel sheets
        
         //read excel file
                XSSFWorkbook wb = new XSSFWorkbook("S:\\Computing\\TAS\\TAS-Research Projects.xlsx");
                
                for(int i = 10; i < 30; i++){
                    
                    BasicDBObject document = new BasicDBObject();
                    BasicDBObject tdoc = new BasicDBObject();
                    
                    Cell length = wb.getSheetAt(0).getRow(i).getCell(5, xc.CREATE_NULL_AS_BLANK);
                    Double lng = 0.0;
                    lng = Null.nullDouble(length, lng);
                    
                    Date date = new Date();
                    
                    Cell ctime = wb.getSheetAt(0).getRow(i).getCell(5, xc.CREATE_NULL_AS_BLANK);
                    Date time = ctime.getDateCellValue();
                    
                    Cell cStart = wb.getSheetAt(0).getRow(i).getCell(3);
                    Date pStart = cStart.getDateCellValue();
                    
                    Cell cEnd = wb.getSheetAt(0).getRow(i).getCell(4);
                    Date pEnd = cEnd.getDateCellValue();
                    
                    boolean between = false;
                    
                    if(pStart != null && pEnd != null){
                        if(date.after(pStart) && date.before(pEnd)){
                            between = true;
                            System.out.println(pStart);
                            System.out.println(pEnd);
                        }
                        else{
                            between = false;
                        }
                    }
                    if(between == false){
                        continue;
                    }
                    else{
                    System.out.println(in);
                    
                    Cell projID = wb.getSheetAt(0).getRow(i).getCell(1, xc.CREATE_NULL_AS_BLANK);
                    String proj = new String();
                    proj = Null.nullString(projID, proj);
                    document.put("uID", proj);
                    
                    document.put("length", lng);
                    
                    Cell cProjectLead = wb.getSheetAt(0).getRow(i).getCell(7, xc.CREATE_NULL_AS_BLANK);
                    String projectLead = new String();
                    projectLead = Null.nullString(cProjectLead, projectLead);
                    String projLead = "";
                    projLead = projectLead;
                    String[] project = new String[4];
                    String[] projectp = projLead.split("\\s+");
                    ArrayList projectL = new ArrayList<String>();          
                    projLead = projLead.toUpperCase();
                    System.out.println(Arrays.toString(projectp));
                    projectL.addAll(Arrays.asList(projectp));
                    
                    if(projectp.length == 2){
                        projectL.add("");
                    }
                    
                    else if(projectp.length == 1){
                         projectL.add("");  
                         projectL.add("");  
                    }
                    document.put("projLead", projectL);
                    
                    
                    Cell cCoinv = wb.getSheetAt(0).getRow(i).getCell(10, xc.CREATE_NULL_AS_BLANK);
                    String coinv = new String();
                    coinv = Null.nullString(cCoinv, coinv);
                    String coinID = "";
                    coinID = coinv;
                    String[] coin1 = new String[4];
                    String[] coin1p = coinv.split("\\s+");
                    ArrayList coINV = new ArrayList<String>();
                    coinID = coinID.toUpperCase();
                    System.out.println(Arrays.toString(coin1p));
                    coINV.addAll(Arrays.asList(coin1p)); 
                    if(coin1p.length == 2){
                        coINV.add("");
                    }
                    else if(coin1p.length == 1){
                        coINV.add("");
                        coINV.add("");
                    }
                    document.put("co inv 1", coINV);
                    
                    
                    Cell cCoinv2 = wb.getSheetAt(0).getRow(i).getCell(13, xc.CREATE_NULL_AS_BLANK);
                    String coinv2 = new String();
                    coinv2 = Null.nullString(cCoinv2, coinv2);
                    String coinID2 = "";
                    coinID2 = coinv2;
                    String[] coin2 = new String[4];
                    String[] coin2p = coinv2.split("\\s+");
                    ArrayList coINV1 = new ArrayList<String>();
                    coinID2 = coinID2.toUpperCase();
                    System.out.println(Arrays.toString(coin2p));
                    coINV1.addAll(Arrays.asList(coin2p));
                    if(coin2p.length == 2){
                        coINV1.add("");
                    }
                    else if(coin2p.length == 1){
                        coINV1.add("");
                        coINV1.add("");
                    }
                    document.put("co inv 2", coINV1);
                    
                   
                    Cell fte = wb.getSheetAt(0).getRow(i).getCell(9, xc.CREATE_NULL_AS_BLANK);
                    Double ft = 0.0;
                    Double fte10 = 0.0;
                    ft = fte.getNumericCellValue();

                                         
                    Cell cfte2 = wb.getSheetAt(0).getRow(i).getCell(12, xc.CREATE_NULL_AS_BLANK);
                    Double fte2 = 0.0;
                    Double fte11 = 0.0;
                    fte2 = cfte2.getNumericCellValue();

                                             
                    Cell cfte3 = wb.getSheetAt(0).getRow(i).getCell(15, xc.CREATE_NULL_AS_BLANK);
                    Double fte3 = 0.0;
                    Double fte12 = 0.0;
                    fte3 = cfte3.getNumericCellValue();


                    Cell cCategory = wb.getSheetAt(0).getRow(i).getCell(6, xc.CREATE_NULL_AS_BLANK);
                    String category = new String();
                    category = Null.nullString(cCategory, category);
                    String[] cate = category.split(",");
                    ArrayList array = new ArrayList<String>();
                    for(int x = 0; x< cate.length; x++){
                        System.out.println(cate[x] + "\n");
                        
                        String[] splitC = category.split("\\s+");
                        ArrayList FinArray = new ArrayList<String>();
                        
                        if(cate[x].contains("Research Councils") || cate[x].contains("research councils") || cate[x].contains("RESEARCH COUNCILS")){
                            array.add(cate[x]);   
                        }
                        else if(cate[x].contains("UK Govt Depts") || cate[x].contains("uk govt depts") || cate[x].contains("UK GOVT DEPTS") || cate[x].contains("UK Government Departments") || cate[x].contains("uk government departments") || cate[x].contains("UK GOVERNMENT DEPARTMENTS")){
                            array.add(cate[x]);
                        }
                        else if(cate[x].contains("European Commission") || cate[x].contains("european commission") || cate[x].contains("EUROPEAN COMMISSION") || cate[x].contains("EU Commission") || cate[x].contains("eu commission") || cate[x].contains("EU COMMISSION")){
                            array.add(cate[x]);
                        }
                        else if(cate[x].contains("UK Charities") || cate[x].contains("uk charities") || cate[x].contains("UK CHARITIES")){
                            array.add(cate[x]);
                        }
                        else if(cate[x].contains("UK Industry") || cate[x].contains("uk industry") || cate[x].contains("UK INDUSTRY")){
                            array.add(cate[x]);
                        }
                        else if(cate[x].contains("KTPs") || cate[x].contains("ktps") || cate[x].contains("KTPS") || cate[x].contains("KTP") || cate[x].contains("ktp")){
                            array.add(cate[x]);
                        }
                        else if(cate[x].contains("Other") || cate[x].contains("OTHER") || cate[x].contains("other")){
                            array.add(cate[x]);
                        }
                        else if(cate[x].contains("PGR") || cate[x].contains("pgr")){
                            array.add(cate[x]);
                        }
                        else if(cate[x].contains("SFC") || cate[x].contains("sfc") || cate[x].contains("Sfc")){
                            array.add(cate[x]);
                        }
                        else if(cate[x].contains("Internal") || cate[x].contains("INTERNAL") || cate[x].contains("internal")){
                            array.add(cate[x]);
                        }
                        else{
                            array.add("");
                        }
                        
                        
                        Matcher m = Pattern.compile("\\((.*?)\\)").matcher(cate[x]);
                        while(m.find() == true) {
                        String per = m.group(0);
                        char charAtZero = per.charAt(1);
                        char charAtOne = per.charAt(2);
                        addChar = "" + charAtZero + charAtOne;
                        charper = Double.parseDouble(addChar);
                        percent = (double)charper/100;
                        }
                        
                        fte10 = ft*percent;
                        array.add(fte10*100);
                        fte11 = fte2*percent;
                        array.add(fte11*100);
                        fte12 = fte3*percent;
                        array.add(fte12*100);
                        
                        if(cate.length <= 1){
                            array.add("");
                        }

                        
                        
                    }
                    if(cate.length <2){
                        for(int a = 1; a<=3; a++){
                            array.add(0);
                        }
                    }
                    
                    
                    document.put("fte", ft*100);
                    document.put("fte2", fte2*100);
                    document.put("fte3", fte3*100);

                    document.put("cate", array);
                    
                    
                    
                    System.out.println("-----------------------------------------");
                    collection.insert(document);
                    in++;

                    
                }
                
                
                
                    
                }
                
    
                
    }
}

                
                
                  
                
                
                 
                    
                    
                    


