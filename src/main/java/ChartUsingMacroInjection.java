import com.aspose.cells.SaveFormat;
import com.aspose.cells.VbaModule;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;

public class ChartUsingMacroInjection {
    public static String readScript() throws IOException {
        File file = new File(
                "./src/main/java/script.vbs");

        BufferedReader br
                = new BufferedReader(new FileReader(file));

        // Declaring a string variable
        String st;
        StringBuilder stringBuilder  = new StringBuilder();
        // Consition holds true till
        // there is character in a string
        while ((st = br.readLine()) != null)
            // Print the string
            stringBuilder.append(st +"\n");

        return stringBuilder.toString();
    }
    public static void main(String[] args) throws Exception {
         Workbook workbook = new Workbook("./src/main/java/workbook.xlsm");
            Worksheet worksheet = workbook.getWorksheets().get(0);
            int idx = workbook.getVbaProject().getModules().add(worksheet);
            VbaModule module = workbook.getVbaProject().getModules().get(idx);
            module.setName("TestModule");
            module.setCodes(readScript());
            workbook.save("output.xlsm", SaveFormat.XLSM);
    }
}
