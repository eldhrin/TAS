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
public class Research {
    
    public static float nullFloat(Cell c, float fl){
        if(c == null){
            fl = 0.0f;
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
    
    public static void research() throws IOException, ParseException{
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
                XSSFWorkbook wb = new XSSFWorkbook("TAS.xlsx");
                
                for(int i = 1; i < 28; i++){
                    BasicDBObject document = new BasicDBObject();
                    BasicDBObject tdoc = new BasicDBObject();
                    
                    Cell cProjectLead = wb.getSheetAt(0).getRow(i).getCell(7, xc.CREATE_NULL_AS_BLANK);
                    String projectLead = new String();
                    projectLead = Null.nullString(cProjectLead, projectLead);
                    String[] plead = projectLead.split(",");
                    String projLead = plead[0];
                    projLead = projLead.toUpperCase();
                    
                    Cell cCoinv = wb.getSheetAt(0).getRow(i).getCell(9, xc.CREATE_NULL_AS_BLANK);
                    String coinv = new String();
                    coinv = Null.nullString(cCoinv, coinv);
                    String[] coin = coinv.split(",");
                    String coinID = coin[0];
                    
                    Cell cCoinv2 = wb.getSheetAt(0).getRow(i).getCell(11, xc.CREATE_NULL_AS_BLANK);
                    String coinv2 = new String();
                    coinv2 = Null.nullString(cCoinv2, coinv2);
                    String[] coin2 = coinv2.split(",");
                    String coinID2 = coin2[0];
                    
                    BasicDBObject match_id = new BasicDBObject("uID", projLead);
                    DBCursor matchCursor = collection.find(match_id);
                    
                    BasicDBObject match_co = new BasicDBObject("uID", coinID);
                    DBCursor matchCo = collection.find(match_co);
                    
                    BasicDBObject match_co2 = new BasicDBObject("uID", coinID2);
                    DBCursor matchCo2 = collection.find(match_co2);
                    
                    if(matchCursor.hasNext() || matchCo.hasNext() || matchCo2.hasNext()){
                        System.out.println("MATCH");
                        document.put("match", "yes");
//                        System.out.println(matchCursor.next());
//                        System.out.println(matchCo.next());
//                        System.out.println(matchCo2.next());
                    
                    
                    Cell cprojectID = wb.getSheetAt(0).getRow(i).getCell(1, xc.CREATE_NULL_AS_BLANK);
                    String projectID = new String();
                    projectID = Null.nullString(cprojectID, projectID);
                    
                    
//                    Cell ctime = wb.getSheetAt(0).getRow(i).getCell(5, xc.CREATE_NULL_AS_BLANK);
//                    Double time = 0.0;
//                    time = Null.nullDouble(ctime, time);
//                    System.out.println(time);
                    
//                    Cell cStart = wb.getSheetAt(0).getRow(i).getCell(4);
//                    Date pStart = new Date();
//                    String sStart = new String();      
//                    pStart = nullDate(cStart, sStart);
//                    System.out.println(pStart);
//                    
//                    Cell cEnd = wb.getSheetAt(0).getRow(i).getCell(5);
//                    Date pEnd = new Date();
//                    String sEnd = new String();
//                    pEnd = nullDate(cEnd, sEnd);
//                    System.out.println(pEnd);
                    

                    Cell cCategory = wb.getSheetAt(0).getRow(i).getCell(6, xc.CREATE_NULL_AS_BLANK);
                    String category = new String();
                    category = Null.nullString(cCategory, category);
                    String[] cate = category.split(",");
                    for(int x = 0; x< cate.length; x++){
                        System.out.println(cate[x] + "\n");
                        if(cate[x].contains("Research Councils") || cate[x].contains("research councils") || cate[x].contains("RESEARCH COUNCILS")){
                            String research_c = cate[x];
                        }
                        else if(cate[x].contains("UK Govt Depts") || cate[x].contains("uk govt depts") || cate[x].contains("UK GOVT DEPTS") || cate[x].contains("UK Government Departments") || cate[x].contains("uk government departments") || cate[x].contains("UK GOVERNMENT DEPARTMENTS")){
                            String uk_gov = cate[x];
                            document.put("uk gov", uk_gov);
                        }
                        else if(cate[x].contains("European Commission") || cate[x].contains("european commission") || cate[x].contains("EUROPEAN COMMISSION") || cate[x].contains("EU Commission") || cate[x].contains("eu commission") || cate[x].contains("EU COMMISSION")){
                            String eu = cate[x];
                            document.put("eu", eu);
                        }
                        else if(cate[x].contains("UK Charities") || cate[x].contains("uk charities") || cate[x].contains("UK CHARITIES")){
                            String uk_charity = cate[x];
                            document.put("uk charity", uk_charity);
                        }
                        else if(cate[x].contains("UK Industry") || cate[x].contains("uk industry") || cate[x].contains("UK INDUSTRY")){
                            String uk_ind = cate[x];
                            document.put("uk ind", uk_ind);
                        }
                        else if(cate[x].contains("KTPs") || cate[x].contains("ktps") || cate[x].contains("KTPS")){
                            String ktp = cate[x];
                            document.put("ktp", ktp);
                        }
                        else if(cate[x].contains("Other") || cate[x].contains("OTHER") || cate[x].contains("other")){
                            String other = cate[x];
                            document.put("other", other);
                        }
                        else if(cate[x].contains("PGR") || cate[x].contains("pgr")){
                            String pgr = cate[x];
                            document.put("pgr", pgr);
                        }
                        else if(cate[x].contains("SFC") || cate[x].contains("sfc") || cate[x].contains("Sfc")){
                            String sfc = cate[x];
                            document.put("sfc", sfc);
                        }
                        else{
                            String other_o = cate[x];
                            document.put("unident", other_o);
                        }
                    }
                    
                    Cell cfte = wb.getSheetAt(0).getRow(i).getCell(8, xc.CREATE_NULL_AS_BLANK);
                    Float fte = 0.0f;
                    String ft = new String();
                    fte = nullFloat(cfte, fte);
                    ft = fte.toString();   
                    
                    Cell cfte2 = wb.getSheetAt(0).getRow(i).getCell(10, xc.CREATE_NULL_AS_BLANK);
                    Float fte2 = 0.0f;
                    String ft2 = new String();
                    fte2 = nullFloat(cfte2, fte2);
                    ft2 = fte2.toString();         
                          
                    Cell cfte3 = wb.getSheetAt(0).getRow(i).getCell(12, xc.CREATE_NULL_AS_BLANK);
                    Float fte3 = 0.0f;
                    String ft3 = new String();
                    fte3 = nullFloat(cfte3, fte3);
                    ft3 = fte3.toString();
                    System.out.println("-----------------------------------------");
                    
                
                    document.put("pID", projectID);
                    //document.put("time", time);
                    document.put("category", cate);
                    document.put("pLead", projLead);
                    document.put("fte", ft);
                    document.put("co-inv", coinID);
                    document.put("fteCO", ft2);
                    document.put("co-inv2", coinID2);
                    document.put("fteCO2", ft3);

                    
                    collection.insert(document);
                    }
                }     
                
            }
                
    }

                
                
                  
                
                
                 
                    
                    
                    


