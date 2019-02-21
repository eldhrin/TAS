package tas;

import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.Mongo;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JOptionPane;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author fl8328
 */
public class DBtoXLSX{
   static int period;
    
    public static void DBtoXLSX() throws IOException, InvalidFormatException{
    Mongo mongo = new Mongo("localhost", 27017);
        DB db = mongo.getDB("TAS");
        //find collection TAS
        DBCollection collection = db.getCollection("TAS");
        int dbcount = (int)collection.count();
        DBCursor cursor = collection.find(); 
        XSSFWorkbook wb = new XSSFWorkbook("H:\\NetBeansProjects\\TAS\\tas_blank.xlsx");
        FileOutputStream fileOut = null;
        
        for(int i = 0; i < dbcount; i++){
            DBObject o = cursor.next();
            DBObject t = (DBObject) o.get("Teaching");
            DBObject r = (DBObject) o.get("Research");
            DBObject s = (DBObject) o.get("Scholarship");
            DBObject q = (DBObject) o.get("Other");
            String name = o.get("name").toString();
            fileOut = new FileOutputStream("H:\\NetBeansProjects\\TAS\\test\\" + name + ".xlsx");
         //TRY CATCH
            //get selected file
                
                //GET VARIABLES FROM THE SPREADSHEET AND CONVERT TO STRING/DOUBLE
                //get name, school, date
                
                Cell cDate = wb.getSheetAt(0).getRow(8).getCell(1);
                String time = o.get("date").toString();
                cDate.setCellValue(time);
                
                Cell cName = wb.getSheetAt(0).getRow(10).getCell(1);
                cName.setCellValue(name);
                
                Cell cuid = wb.getSheetAt(0).getRow(11).getCell(1);
                String uid = o.get("uID").toString();
                cuid.setCellValue(uid);
                
                Cell cSchool = wb.getSheetAt(0).getRow(12).getCell(1);
                String cschool = o.get("school").toString();
                cSchool.setCellValue(cschool);
                      
                //TEACHING
                Cell cCore = wb.getSheetAt(0).getRow(16).getCell(2);
                Cell cCores = wb.getSheetAt(0).getRow(43).getCell(5);
                Double core = Double.parseDouble(t.get("core").toString());
                cCore.setCellValue(core);
                cCores.setCellValue(core);

                Cell cSupport = wb.getSheetAt(0).getRow(17).getCell(2);
                Cell cSupports = wb.getSheetAt(0).getRow(43).getCell(6);
                Double support = Double.parseDouble(t.get("support").toString());
                cSupport.setCellValue(support);
                cSupports.setCellValue(support);
                
                //RESEARCH
                Cell cCouncils = wb.getSheetAt(0).getRow(20).getCell(2);
                Cell cCouncilss = wb.getSheetAt(0).getRow(43).getCell(7);
                Double councils = Double.parseDouble(r.get("council").toString());
                cCouncils.setCellValue(councils);
                cCouncilss.setCellValue(councils);
//                
//                Cell cUK_govt = wb.getSheetAt(0).getRow(21).getCell(2);
//                Cell cUK_govts = wb.getSheelAt(0).getRow(43).getCell(8);
//                Double UK_govt = Double.parseDouble(r.get("UK_govt").toString());
//                cUK_govt.setCellValue(UK_govt);
//                cUK_govts.getCellValue(UK_govt);
//                
//                Cell cEU = wb.getSheetAt(0).getRow(22).getCell(2);
//                Cell cEUs = wb.getSheetAt(0).getRow(43).getCell(9);
//                Double EU = Double.parseDouble(r.get("EU").toString());
//                cEU.setCellValue(EU);
//                cEUs.setCellValue(EU);
//                
//                Cell cUK_charity = wb.getSheetAt(0).getRow(23).getCell(2);   
//                Cell cUK_charitys = wb.getSheetAt(0).getRow(43).getCell(10);
//                Double UK_charity = Double.parseDouble(r.get("UK_charity").toString());
//                cUK_charity.setCellValue(UK_charity);
//                cUK_charitys.setCellValue(UK_charitys);
//                
//                Cell cUK_industry = wb.getSheetAt(0).getRow(24).getCell(2);
//                Cell cUK_industrys = wb.getSheetAt(0).getRow(43).getCell(11);
//                Double UK_industry = Double.parseDouble(r.get("UK_industry").toString());
//                cUK_industry.setCellValue(UK_industry);
//                cUK_industrys.setCellValue(UK_industry);
//                
//                Cell cKTP_projects = wb.getSheetAt(0).getRow(25).getCell(2); 
//                Cell cKTP_projetctss = wb.getSheetAt(0).getRow(43).getCell(12);
//                Double KTP_projects = Double.parseDouble(r.get("KTP_projects").toString());
//                cKTP_projects.setCellValue(KTP_projects);
//                cKTP_projectss.setCellValue(KTP_projects);
//                
//                Cell cOther = wb.getSheetAt(0).getRow(26).getCell(2);
//                Cell cOthers = wb.getSheetAt(0).getRow(43).getCell(13);
//                Double other = Double.parseDouble(r.get("other").toString());
//                cOther.setCellValue(other);
//                cOthers.setCellValue(other);
//                
//                Cell cSFC_innovation = wb.getSheetAt(0).getRow(27).getCell(2);
//                Cell cSFC_innovations = wb.getSheetAt(0).getRow(43).getCell(14);
//                Double SFC_innovation = Double.parseDouble(r.get("SFC_innovation").toString());
//                cSFC_innovation.setCellValue(SFC_innovation);
//                cSFC_innovations.setCellValue(SFC_innovation);
//                
//                Cell cSFC_RD = wb.getSheetAt(0).getRow(28).getCell(2);
//                Cell cSFC_RDs = wb.getSheetAt(0).getRow(43).getCell(15);
//                Double SFC_RD = Double.parseDouble(r.get("SFC_RD").toString());
//                cSFC_RD.setCellValue(SFC_RD);
//                cSFC_RDs.setCellValue(SFC_RD);
//                
//                Cell cPGR_supervision = wb.getSheetAt(0).getRow(29).getCell(2);     
//                Cell cPGR_supervisions = wb.getSheetAt(0).getRow(43).getCell(16);
//                Double PGR_supervision = Double.parseDouble(r.get("PGR_supervision").toString());
//                cPGR_supervision.setCellValue(PGR_supervision);
//                cPGR_supervisions.setCellValue(PGR_supervision);
//                
//                Cell cInternal_research = wb.getSheetAt(0).getRow(30).getCell(2); 
//                Cell cInternal_researchs = wb.getSheetAt(0).getRow(43).getCell(17);
//                Double internal_research = Double.parseDouble(r.get("internal_research").toString());
//                cInternal_research.setCellValue(internal_research);
//                cInternal_researchs.setCellValue(internal_research);
//               
//                Cell cSupport_intext= wb.getSheetAt(0).getRow(31).getCell(2);
//                Cell cSupport_intexts = wb.getSheetAt(0).getRow(43).getCell(18);
//                Double support_intext = Double.parseDouble(r.get("support_intext").toString());
//                cSupport_intext.setCellValue(support_intext);
//                cSupport_intexts.setCellValue(support_intext);
////                
//                Cell cSupport_SFC = wb.getSheetAt(0).getRow(32).getCell(2);
//                Cell cSupport_SFCs = wb.getSheetAt(0).getRow(43).getCell(19);
//                Double support_SFC = Double.parseDouble(r.get("support_SFC").toString());
//                cSupport_SFC.setCellValue(support_SFC);
//                cSupport_SFCs.setCellValue(support_SFC);
//                
//                //SCHOLARSHIP
//                Cell cTeaching = wb.getSheetAt(0).getRow(34).getCell(2);
//                Cell cTeachings = wb.getSheetAt(0).getRow(43).getCell(20);
//                Double teaching = Double.parseDouble(s.get("teaching").toString());
//                cTeaching.setCellValue(teaching);
//                cTeachings.setCellValue(teaching);
//                
//                Cell cResearch = wb.getSheetAt(0).getRow(35).getCell(2); 
//                Cell cResearchs = wb.getSheetAt(0).getRow(43).getCell(21);
//                Double research = Double.parseDouble(s.get("research").toString());
//                cResearch.setCellValue(research);
//                cResearchs.setCellValue(research);
//                
//                Cell cPhD = wb.getSheetAt(0).getRow(36).getCell(2); 
//                Cell cPhDs = wb.getSheetAt(0).getRow(43).getCell(22);
//                Double PhD = Double.parseDouble(s.get("PhD").toString());
//                cPhD.setCellValue(PhD);
//                cPhDs.setCellValue(PhD);
                
                //OTHER
                Cell coOther = wb.getSheetAt(0).getRow(38).getCell(2);
                Cell coOthers = wb.getSheetAt(0).getRow(43).getCell(23);
                Double cother = Double.parseDouble(q.get("Other").toString());
                coOther.setCellValue(cother);
                coOthers.setCellValue(cother);
                
//                Cell coSupport = wb.getSheetAt(0).getRow(39).getCell(2);
//                Cell coSupports = wb.getSheetAt(0).getRow(43).getCell(24);
//                Double cosupport = Double.parseDouble(q.get("Osupport").toString());
//                coSupport.setCellValue(cosupport);
//                coSupports.setCellValue(cosupport);
//                
//                //MANAGEMENT
                Cell cMgmt = wb.getSheetAt(0).getRow(41).getCell(2);      
                Cell cMgmts = wb.getSheetAt(0).getRow(43).getCell(25);
                Double mgmt = Double.parseDouble(o.get("Mgmt").toString());
                cMgmt.setCellValue(mgmt);
                cMgmts.setCellValue(mgmt);
                
//                //TOTAL
//                Double ctotal = core + support + councils + UK_govt + EU + UK_charity + UK_industry + KTP_projects + other + SFC_innovation + SFC_RD + PGR_supervision + internal_research + support_intext + support_SFC + teaching + research + PhD + cother + cosupport + mgmt;
//                Cell cTotal = wb.getSheetAt(0).getRow(43).getCell(2);
//                cTotal.setCellValue(ctotal);
                Cell ctotal = wb.getSheetAt(0).getRow(43).getCell(2);
                Cell cTotals = wb.getSheetAt(0).getRow(43).getCell(26);
                Double total = core+support+mgmt+councils+cother;
                ctotal.setCellValue(total);
                cTotals.setCellValue(total);
                
                Cell cHols = wb.getSheetAt(0).getRow(45).getCell(2);   

                //HOLIDAYS
                int h = Null.getDate();
        switch (h) {               
            case 1:
                {
                    Double hols = Double.parseDouble(o.get("sem3").toString());
                    cHols.setCellValue(hols);
                    break;
                }
            case 2:          
                {
                    Double hols = Double.parseDouble(o.get("sem1").toString());
                    cHols.setCellValue(hols);
                    break;
                }
            default:              
                {
                    Double hols = Double.parseDouble(o.get("sem2").toString());
                    cHols.setCellValue(hols);
                    break;
                }
        }
                
                
                System.out.println("Data from Database to excel file");
                wb.write(fileOut);
        }
        JOptionPane.showMessageDialog(null, "Done! \n Reports are saved in the test folder", "Info", JOptionPane.INFORMATION_MESSAGE);
        wb.close();
        mongo.close();
    }
}
