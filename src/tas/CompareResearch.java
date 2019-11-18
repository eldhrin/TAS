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
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 *
 * @author fl8328
 */
public class CompareResearch {
    
    
    public static void compareResearch() throws IOException, InvalidFormatException, ParseException{
     Mongo mongo = new Mongo("localhost", 27017);
     DB tas = mongo.getDB("TAS");
     DB research = mongo.getDB("RESEARCH");
     //find collection TAS
     DBCollection tCollection = research.getCollection("RESEARCH");
     DBCollection tCollR = tas.getCollection("NEWTAS");
     
     int dbcount = (int)tCollR.count();
     int db = (int)tCollection.count();
     
     
     //find one from main DB, get the ID and match to research db
        DBCursor cursor = tCollR.find();
        DBCursor compare = tCollection.find();
     
     for(int i = 1; i < dbcount; i++){
         
         
        DBObject o = cursor.next();
        
            
        Double fte = 0.0;
        Double fte2 = 0.0;
        Double fte3 = 0.0;
         System.out.println("here");

            if(compare.hasNext()){
                String id = o.get("id").toString();
                DBObject c = compare.next();
                String proID = c.get("projLead").toString();
                String coInv = c.get("co inv 1").toString();
                String coInv2 = c.get("co inv 2").toString();

                if(id.equals(proID)){
                    fte += Double.parseDouble(c.get("fte").toString());
                    System.out.println(fte);
                    System.out.println("found1");
                }
                else if(id.equals(coInv)){
                    fte2 += Double.parseDouble(c.get("fte2").toString());
                    System.out.println(fte2);
                    System.out.println("found2");
                }
                else if(id.equals(coInv2)){
                    fte3 += Double.parseDouble(c.get("fte3").toString());
                    System.out.println(fte3);
                    System.out.println("found3");
                }
                else{
                    System.out.println("not found");
                    

                }
            }
        }

    
    }
}

