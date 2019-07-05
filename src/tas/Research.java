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
import java.util.Date;
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
        DB tas = mongo.getDB("RESEARCH");
        //find collection TAS
        DBCollection collection = tas.getCollection("RESEARCH");
        
        //remove all entries from the database
        DBCursor rem = collection.find();
        while(rem.hasNext()){
            collection.remove(rem.next());
        }
        int in = 0;
        //user chooses directory containing all users tas excel sheets
        
         //read excel file
                XSSFWorkbook wb = new XSSFWorkbook("S:\\Computing\\TAS\\TAS-Research Projects.xlsx");
                
                for(int i = 5; i < 26; i++){
                    System.out.println(in);
                    BasicDBObject document = new BasicDBObject();
                    BasicDBObject tdoc = new BasicDBObject();
                    
                    Cell projID = wb.getSheetAt(0).getRow(i).getCell(1, xc.CREATE_NULL_AS_BLANK);
                    String proj = new String();
                    proj = Null.nullString(projID, proj);
                    document.put("uID", proj);
                    
                    Cell cProjectLead = wb.getSheetAt(0).getRow(i).getCell(7, xc.CREATE_NULL_AS_BLANK);
                    String projectLead = new String();
                    projectLead = Null.nullString(cProjectLead, projectLead);
                    String[] plead = projectLead.split(",");
                    String projLead = plead[0];
                    projLead = projLead.toUpperCase();
                    document.put("projLead", projLead);
                    
                    Cell cCoinv = wb.getSheetAt(0).getRow(i).getCell(10, xc.CREATE_NULL_AS_BLANK);
                    String coinv = new String();
                    coinv = Null.nullString(cCoinv, coinv);
                    String[] coin = coinv.split(",");
                    String coinID = coin[0];
                    coinID = coinID.toUpperCase();
                    document.put("co inv 1", coinID);
                    
                    Cell cCoinv2 = wb.getSheetAt(0).getRow(i).getCell(13, xc.CREATE_NULL_AS_BLANK);
                    String coinv2 = new String();
                    coinv2 = Null.nullString(cCoinv2, coinv2);
                    String[] coin2 = coinv2.split(",");
                    String coinID2 = coin2[0];
                    coinID2 = coinID2.toUpperCase();
                    document.put("co inv 2", coinID2);
                    
                    Cell length = wb.getSheetAt(0).getRow(i).getCell(5, xc.CREATE_NULL_AS_BLANK);
                    Double lng = 0.0;
                    lng = Null.nullDouble(length, lng);
                   
                    Cell fte = wb.getSheetAt(0).getRow(i).getCell(9, xc.CREATE_NULL_AS_BLANK);
                    Double ft = 0.0;
                    System.out.println("evaluated as " + fte.getNumericCellValue());
                    ft = fte.getNumericCellValue();
                    ft = ft*lng;
                    ft = ft/4;
                    document.put("fte", ft);
                    
                    Date date = new Date();
                    System.out.println(date);
                    
                    Cell ctime = wb.getSheetAt(0).getRow(i).getCell(5, xc.CREATE_NULL_AS_BLANK);
                    Date time = ctime.getDateCellValue();
                    System.out.println(time);
                    
                    Cell cStart = wb.getSheetAt(0).getRow(i).getCell(3);
                    Date pStart = cStart.getDateCellValue();
                    System.out.println(pStart);
                    
                    Cell cEnd = wb.getSheetAt(0).getRow(i).getCell(4);
                    Date pEnd = cEnd.getDateCellValue();
                    System.out.println(pEnd);
                    
                    boolean between = false;
                    
                    if(date != null && pStart != null && pEnd != null){
                        if(date.after(pStart) && date.before(pEnd)){
                            between = true;
                        }
                        else{
                            between = false;
                        }
                    }
                    System.out.println(between);

                    Cell cCategory = wb.getSheetAt(0).getRow(i).getCell(6, xc.CREATE_NULL_AS_BLANK);
                    String category = new String();
                    category = Null.nullString(cCategory, category);
                    String[] cate = category.split(",");
                    ArrayList array = new ArrayList<String>();
                    for(int x = 0; x< cate.length; x++){
                        System.out.println(cate[x] + "\n");
                        
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
                        else if(cate[x].contains("KTPs") || cate[x].contains("ktps") || cate[x].contains("KTPS")){
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
                        else{
                            array.add(cate[x]);
                        }
                        
                    }
                    document.put("cate", array);
                                         
                    Cell cfte2 = wb.getSheetAt(0).getRow(i).getCell(12, xc.CREATE_NULL_AS_BLANK);
                    Double fte2 = 0.0;
                    fte2 = cfte2.getNumericCellValue();
                    fte2 = fte2*lng;
                    fte2 = fte2/4;
                    document.put("fte2", fte2);     
                    
                          
                    Cell cfte3 = wb.getSheetAt(0).getRow(i).getCell(15, xc.CREATE_NULL_AS_BLANK);
                    Double fte3 = 0.0;
                    fte3 = cfte3.getNumericCellValue();
                    fte3 = fte3*lng;
                    fte3 = fte3/4;
                    document.put("fte3", fte3);
                    System.out.println("-----------------------------------------");
                    if(between == true){
                        collection.insert(document);
                    }
                    in++;

                    
                }
                
                
                
        }     
                
    
                
    }

                
                
                  
                
                
                 
                    
                    
                    


