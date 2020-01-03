import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class POISearch {
    public static <Row> void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream(new File("gfgcontribute.xlsx"));

            // Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Search for Tony
            for (org.apache.poi.ss.usermodel.Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        if (cell.getRichStringCellValue().getString().trim().equals("Tony")) {
                            System.out.println(row.getRowNum());
                            break;
                        }
                    }
                }
            }

            file.close();


        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }
}
