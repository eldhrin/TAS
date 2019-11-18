//Adam Lyons 20/11/2018
// filter to read xls, xlsx files
//reads multiple files in the folder with those extensions
package tas;

import java.io.File;

/**
 *
 * @author fl8328
 */
    public class ExcelFileFilter implements java.io.FileFilter{

    @Override
    //only accepts certain filetypes
    public boolean accept(File file) {
        return file != null &&
                file.isFile() &&
                file.canRead() &&
                //legacy excel files
                (file.getName().endsWith("xls") ||
                 file.getName().endsWith("XLS") ||
                //modern excel files
                 file.getName().endsWith("xlsx")||
                 file.getName().endsWith("XLSX"));
    }
        
}
