//Adam Lyons 29/01/2019
package tas;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.Mongo;
import java.io.File;
import java.io.IOException;
import java.io.StringWriter;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class ReadXml {
    
    
    public static void readXml() throws ParserConfigurationException, SAXException, IOException, TransformerException{
        
        //connect to local mongodb
        Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("TAS");
        
         int dbcount = (int)collection.count();
//         
//         for(int i = 0; i < dbcount; i++){
//            DBObject o = cursor.next();
//            DBObject t = (DBObject) o.get("Teaching");
//            DBObject r = (DBObject) o.get("Research");
//            DBObject s = (DBObject) o.get("Scholarship");
//            DBObject q = (DBObject) o.get("Other");
//         

	File file = new File("worktribe.xml");
        
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        DocumentBuilder docb = dbf.newDocumentBuilder();
        Document doc = docb.parse(file);
        
        doc.getDocumentElement().normalize();
        System.out.println("Root element of the doc is " + doc.getDocumentElement().getNodeName());
        
        NodeList lof = doc.getElementsByTagName("Person");
        int total = lof.getLength();
        System.out.println("Total no of people "+ total);
        
        NodeList lofp = doc.getElementsByTagName("Project");
        int pro = lofp.getLength();
        System.out.println("Total no of projects: " + pro);
        
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        
        for(int j = 0; j < lofp.getLength(); j++){
            Node fpn = lof.item(j);
            if(fpn.getNodeType() == fpn.ELEMENT_NODE){
                
            Element eElement = (Element) fpn;
            DOMSource source = new DOMSource(fpn);
            StreamResult result = new StreamResult(new StringWriter());

            transformer.transform(source,result);
            String id = eElement.getElementsByTagName("Login").item(0).getTextContent();
            
            BasicDBObject query = new BasicDBObject();
            query.put("uID", id);
            DBCursor cursor = collection.find(query);
            
            while(cursor.hasNext()){
                Node pn = lofp.item(j);
                    Element eProject = (Element) pn;
                    DOMSource pSource = new DOMSource(pn);
                    StreamResult proj = new StreamResult(new StringWriter());

                    transformer.transform(pSource, proj);
            }
                System.out.println(id);
        }
//        
//        for(int i = 0; i<lof.getLength(); i++){
//            
//            Node fpn = lof.item(i);
//            if(fpn.getNodeType() == Node.ELEMENT_NODE){
//                
//                Element fpe = (Element)fpn;
//                
//                NodeList firstname = fpe.getElementsByTagName("Person");
//                Element name = (Element)firstname.item(0);
//                
//                NodeList textList = name.getChildNodes();
//                System.out.println("Name : " + ((Node)textList.item(0)).getNodeValue().trim());
//            
//            }
//        }
  }
    }
}

