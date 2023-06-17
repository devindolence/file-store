package reference.file;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONObject;

import java.util.Iterator;

public class DtoToExcelConverter {
    //
    public static <T extends ExcelDto> JSONObject toJsonObject(T excelDto) {
        ObjectMapper objectMapper = new ObjectMapper();
        String jsonString;

        try {
            jsonString = objectMapper.writeValueAsString(excelDto);
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }

        return new JSONObject(jsonString);
    }

    public Workbook toExcel(
            JSONObject jsonObject
    ) {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Create the headers
        Row headerRow = sheet.createRow(0);
        Row subheaderRow = sheet.createRow(1);
        Iterator<String> keys = jsonObject.keys();
        int columnCount = 0;
        while (keys.hasNext()) {
            String key = keys.next();
            Cell headerCell = headerRow.createCell(columnCount);
            headerCell.setCellValue(key);
            if (jsonObject.get(key) instanceof JSONObject) {
                JSONObject subObject = jsonObject.getJSONObject(key);
                Iterator<String> subKeys = subObject.keys();
                int subColumnCount = columnCount;
                while (subKeys.hasNext()) {
                    String subKey = subKeys.next();
                    Cell subheaderCell = subheaderRow.createCell(subColumnCount);
                    subheaderCell.setCellValue(subKey);
                    subColumnCount++;
                }
            }
            columnCount++;
        }

        // Add data to the cells
        Row dataRow = sheet.createRow(2);
        keys = jsonObject.keys();
        columnCount = 0;
        while (keys.hasNext()) {
            String key = keys.next();
            if (jsonObject.get(key) instanceof JSONObject) {
                JSONObject subObject = jsonObject.getJSONObject(key);
                Iterator<String> subKeys = subObject.keys();
                int subColumnCount = columnCount;
                while (subKeys.hasNext()) {
                    String subKey = subKeys.next();
                    Cell dataCell = dataRow.createCell(subColumnCount);
                    dataCell.setCellValue(subObject.getString(subKey));
                    subColumnCount++;
                }
            } else {
                Cell dataCell = dataRow.createCell(columnCount);
                dataCell.setCellValue(jsonObject.getString(key));
            }
            columnCount++;
        }

        return workbook;
    }
}
