import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class MacroEnabledExcelUsingApachePoi {
    public static void main(String[] args) throws InvalidFormatException, IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open("input.xlsm"));
        XSSFSheet sheet = workbook.getSheetAt(0);

        Object[][] bookData = {
                {"KPI", "Average Patient", "Quality Stars", "Rehab", "SNF", "MHSA","LTACH"},
                {"CYTD", 108, 108, 65, 300,565,null},
                {"PYTD", 96, 96, 89, 455, 400,455},
                {"PFY", 145, 145, 100, 335,545,400}
        };

        int rowCount = 3;

        for (Object[] aBook : bookData) {
            Row row = sheet.createRow(++rowCount);

            int columnCount = 0;

            for (Object field : aBook) {
                Cell cell = row.createCell(++columnCount);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }

        }


        try (FileOutputStream outputStream = new FileOutputStream("output.xlsm")) {
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
