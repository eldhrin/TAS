package tas;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.Mongo;
import com.mongodb.MongoClient;
import com.mongodb.MongoClientURI;
import org.apache.poi.xssf.usermodel.*;

import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import com.mongodb.client.MongoIterable;
import com.mongodb.util.JSON;

import org.json.*;

import java.util.*;
import java.io.*;
import java.net.UnknownHostException;
import javax.swing.*;
import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.bson.Document;
import org.bson.types.ObjectId;


/**
 *
 * @author fl8328
 */
public class TAS {
    
    

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, JSONException {
        
        GetXlsx.getXlsx();
    }
    
}
