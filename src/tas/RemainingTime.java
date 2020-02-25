/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tas;

/**
 *
 * @author fl8328
 */
public class RemainingTime {
    
     adjusted = remTime - totRes;
            if(adjusted < 0.0){
                    System.out.println("ERROR !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
            }
            System.out.println("REMAINING TIME: " + remTime);
            System.out.println("NULL " + adjusted);
                
            Double div = adjusted/100;
                
            Double split45 = div*45;
                
            Double split10 = div*10;
                
            System.out.println("nnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn");
            System.out.println(split45);
            System.out.println(split10);
                

                
            tSupp = Null.nullDouble(tsup, tSupp);
            tSupp = tSupp + split45;
            tsup.setCellValue(tSupp);
                
            scholt = Null.nullDouble(cscholt, scholt);
            scholt = scholt + split45;
            cscholt.setCellValue(scholt);
                
            mgmt = Null.nullDouble(cmgmt, mgmt);
            mgmt = mgmt + split10;
            cmgmt.setCellValue(mgmt);
    
    
     else{
                adjusted = research - totRes;
                if(adjusted < 0.0){
                    System.out.println("ADJUSTED " + research + " - " + totRes + " = " + adjusted);
                    System.out.println("ERROR !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                System.out.println("FULL " + adjusted);
                
                Double div = adjusted/100;
                
                Double split80 = div*80; 
                Double split20 = div*20;
                
                System.out.println(split80);
                System.out.println(split20);
                
                Double ressupport = 0.0;
                ressupport = suppint + split80;
                
                System.out.println(ressupport);
                SuppInt.setCellValue(ressupport); 
                
                phd = phd + split20;
                cphd.setCellValue(phd);
                
                System.out.println("REMAINING TIME: " + remTime);
                Double remtime = remTime/100;
                
                Double split451 = remtime*45;
                Double split101 = remtime*10;
                
                System.out.println(split451);
                System.out.println(split101);
               
                tSupp = tSupp + split451;
                tsup.setCellValue(tSupp);
                
                scholt = scholt + split451;
                cscholt.setCellValue(scholt);
                
                mgmt = mgmt + split101;
                cmgmt.setCellValue(mgmt);
                

            }
}
