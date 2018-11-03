package tas;

import java.io.IOException;
import java.text.ParseException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.json.*;



/**
 *
 * @author fl8328
 */
public class TAS {
    
    

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, InvalidFormatException, ParseException {
        
        //GetXlsx.getXlsx();
        //WriteXlsx.writeXlsx();
        //GenerateXlsx.generateXlsx();
        GetHols.getHols();
    }
    
}
