package tas;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;

import com.mongodb.DB;

import org.json.*;

import java.util.*;
import java.io.*;
import javax.swing.*;


/**
 *
 * @author fl8328
 */
public class TAS {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, JSONException {
       
        //user chooses file
        JFileChooser fileChooser = new JFileChooser();
        int returnValue = fileChooser.showOpenDialog(null);
        //approve file chosen
        if(returnValue == JFileChooser.APPROVE_OPTION){
            
            //TRY CATCH
            //get selected file
            try {
                
                XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
                wb.setMissingCellPolicy(MissingCellPolicy.RETURN_NULL_AND_BLANK);
                
                String date = wb.getSheetAt(0).getRow(8).getCell(1).toString();
                String name = wb.getSheetAt(0).getRow(10).getCell(1).toString();
                String school = wb.getSheetAt(0).getRow(12).getCell(1).toString();
                
                Double core = Double.parseDouble(wb.getSheetAt(0).getRow(16).getCell(2).toString());
                Double support = Double.parseDouble(wb.getSheetAt(0).getRow(17).getCell(2).toString());
                Double councils = Double.parseDouble(wb.getSheetAt(0).getRow(20).getCell(2).toString());
                Double UK_govt = Double.parseDouble(wb.getSheetAt(0).getRow(21).getCell(2).toString());
                Double EU = Double.parseDouble(wb.getSheetAt(0).getRow(22).getCell(2).toString());
                Double UK_charity = Double.parseDouble(wb.getSheetAt(0).getRow(23).getCell(2).toString());
                Double UK_industry = Double.parseDouble(wb.getSheetAt(0).getRow(24).getCell(2).toString());
                Double KTP_projects = Double.parseDouble(wb.getSheetAt(0).getRow(25).getCell(2).toString());
                Double other = Double.parseDouble(wb.getSheetAt(0).getRow(26).getCell(2).toString());
                Double SFC_innovaton = Double.parseDouble(wb.getSheetAt(0).getRow(27).getCell(2).toString());
                Double SFC_RD = Double.parseDouble(wb.getSheetAt(0).getRow(28).getCell(2).toString());
                Double PGR_supervision = Double.parseDouble(wb.getSheetAt(0).getRow(29).getCell(2).toString());
                Double internal_research = Double.parseDouble(wb.getSheetAt(0).getRow(30).getCell(2).toString());
                Double support_intext= Double.parseDouble(wb.getSheetAt(0).getRow(31).getCell(2).toString());
                Double support_SFC = Double.parseDouble(wb.getSheetAt(0).getRow(32).getCell(2).toString());
                
                Double teaching = Double.parseDouble(wb.getSheetAt(0).getRow(34).getCell(2).toString());
                Double research = Double.parseDouble(wb.getSheetAt(0).getRow(35).getCell(2).toString());
                Double PhD = Double.parseDouble(wb.getSheetAt(0).getRow(36).getCell(2).toString());
                
                Double oOther = Double.parseDouble(wb.getSheetAt(38).getRow(17).getCell(2).toString());
                Double oSupport = Double.parseDouble(wb.getSheetAt(39).getRow(17).getCell(2).toString());
                
                Double mgmt = Double.parseDouble(wb.getSheetAt(0).getRow(41).getCell(2).toString());
                
                Double total = Double.parseDouble(wb.getSheetAt(0).getRow(43).getCell(2).toString());
                
                Double hols = Double.parseDouble(wb.getSheetAt(0).getRow(45).getCell(2).toString());
                System.out.println(support + hols + teaching + PhD);
            } catch (FileNotFoundException ex) {
               ex.printStackTrace();           
            }
        }
    }
}
