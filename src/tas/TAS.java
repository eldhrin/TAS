package tas;

import java.io.IOException;
import java.text.ParseException;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.xml.sax.SAXException;



/**
 *
 * @author fl8328
 */
public class TAS {
    
    

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, InvalidFormatException, ParseException, ParserConfigurationException, SAXException, TransformerException {
        //Teaching.teaching();
        //GetHols.getHols();
        //DBtoXLSX.DBtoXLSX();
        //SWLtoDB.SWLtoDB();
        //DBtoXLSX.DBtoXLSX();
        //GetHols.getHols();
        //DBtoXLSX.DBtoXLSX();
        //WriteReport.writeReport();
        //XLSXtoDB.xlsxtoDB();
        //ReadXml.readXml();
        //AwardstoDB.awardstoDB();
        //Research.research();
        //Teaching.teaching();
        //TeachingXLSX.teachingXLSX();
        //CompareResearch.compareResearch();
        //TeachingTemp.teachingTemp();
        
//        GetPGR.GetPGR();
//        // TEST PutPGR.PutPGR();
        NullifyXLSX.nullify();
//        GetWorkloadModel.GetWorkloadModel();
        GrabRemainingTime.GrabRemainingTime(); // WRITES PGR ALLOCATION
        WriteReportWorkload.writeReport();
        WriteReportResearch.writeReport();
//        Research.research();
//        
        CompareResearch.compareResearch();

          //GetHols2.getHols2();
    }
    
}
