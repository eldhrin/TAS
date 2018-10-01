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
    public boolean accept(File file) {
        return file != null &&
                file.isFile() &&
                file.canRead() &&
                (file.getName().endsWith("xls") ||
                 file.getName().endsWith("xlsx"));
    }
        
}
