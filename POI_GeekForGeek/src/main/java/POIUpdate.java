import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class POIUpdate {

    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream(new File("gfgcontribute.xlsx"));

            // Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            XSSFCell updatingCell = sheet.getRow(1).getCell(1);
            if(updatingCell!=null){
                updatingCell.setCellValue("UPDATED");
            }

            file.close();


            // this Writes the workbook gfgcontribute
            FileOutputStream out = new FileOutputStream(new File("output.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("output.xlsx written successfully on disk.");


        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }
}
