package reference;

import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;
import reference.dto.TestUser;
import reference.dto.TestUser2;
import reference.file.DtoToExcelConverter;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {

        List<TestUser> testUsers = new ArrayList<>();
        List<TestUser2> testUser2List = new ArrayList<>();

        for (int i = 0; i < 100; i++) {
            TestUser testUser = new TestUser("test : " + i, String.valueOf(i));
            TestUser testUser2 = new TestUser("test2 : " + i, String.valueOf(i));
            testUsers.add(testUser);
            testUsers.add(testUser2);

            TestUser2 twoHeaderDto = new TestUser2(testUser, testUser2);
            testUser2List.add(twoHeaderDto);
        }

        JSONArray oneHeader = DtoToExcelConverter.toJsonArray(testUsers);
        JSONArray twoHeader = DtoToExcelConverter.toJsonArray(testUser2List);

        // Create a new Excel workbook
        // Create a new sheet in the workbook
        DtoToExcelConverter dtoToExcelConverter = new DtoToExcelConverter();

        Workbook workbook = dtoToExcelConverter.toExcel(oneHeader, testUsers.size());
        Workbook workbook2 = dtoToExcelConverter.toExcel(twoHeader, testUser2List.size());


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
