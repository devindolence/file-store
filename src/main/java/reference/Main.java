package reference;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONObject;
import reference.dto.TestUser;
import reference.dto.TestUser2;
import reference.file.DtoToExcelConverter;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class Main {
    public static void main(String[] args) {

        TestUser testUser = new TestUser("test1", "1");
        TestUser testUser2 = new TestUser("test2", "2");
        TestUser2 twoHeaderDto = new TestUser2(testUser, testUser2);

        JSONObject oneHeader = DtoToExcelConverter.toJsonObject(testUser);
        JSONObject twoHeader = DtoToExcelConverter.toJsonObject(twoHeaderDto);

        // Create a new Excel workbook


        // Create a new sheet in the workbook
        DtoToExcelConverter dtoToExcelConverter = new DtoToExcelConverter();

        Workbook workbook = dtoToExcelConverter.toExcel(oneHeader);
        Workbook workbook2 = dtoToExcelConverter.toExcel(twoHeader);


        // Write to an Excel file
        try (OutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        try (OutputStream fileOut = new FileOutputStream("workbook2.xlsx")) {
            workbook2.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
