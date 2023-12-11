package reference.file;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.util.Collection;
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

    public static <T extends ExcelDto> JSONArray toJsonArray(Collection<T> excelDto) {
        ObjectMapper objectMapper = new ObjectMapper();
        String jsonString;

        try {
            jsonString = objectMapper.writeValueAsString(excelDto);
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }

        return new JSONArray(jsonString);
    }

    public Workbook toExcel(
            JSONObject jsonObject
    ) {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Create the headers
        createHeader(jsonObject, sheet);

        // Add data to the cells
        addCellData(jsonObject, sheet);

        return workbook;
    }

    // list to excel
    public Workbook toExcel(
            JSONArray jsonArray,
            int size
    ) {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Create the headers
        createHeader(jsonArray, sheet);

        // Add data to the cells
        addCellData(jsonArray, sheet, size);

        return workbook;
    }

    // object to cell
    private void addCellData(JSONObject jsonObject, Sheet sheet) {
        int columnCount;
        Iterator<String> keys;
        Row dataRow = sheet.createRow(2);
        keys = jsonObject.keys();
        columnCount = 0;
        int subColumnCount = 0;
        while (keys.hasNext()) {
            String key = keys.next();

            // if exists sub header
            if (jsonObject.get(key) instanceof JSONObject) {
                JSONObject subObject = jsonObject.getJSONObject(key);
                Iterator<String> subKeys = subObject.keys();
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
    }

    // todo refactor
    private void addCellData(JSONArray jsonArray, Sheet sheet, int size) {
        int columnCount;
        for (int i = 0 ; i < size ; i++) {
            Iterator<String> keys;
            Row dataRow = sheet.createRow(2+i);
            JSONObject jsonObject = jsonArray.getJSONObject(i);
            keys = jsonObject.keys();
            columnCount = 0;
            int subColumnCount = 0;
            while (keys.hasNext()) {
                String key = keys.next();

                // if exists sub header
                if (jsonObject.get(key) instanceof JSONObject) {
                    JSONObject subObject = jsonObject.getJSONObject(key);
                    Iterator<String> subKeys = subObject.keys();
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
        }
    }

    // JSONObject > T Property to Sheet Header Column. if T Object has group header, rows setting.
    private void createHeader(JSONObject jsonObject, Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        Row subheaderRow = sheet.createRow(1);
        Iterator<String> keys = jsonObject.keys();
        int columnCount = 0;
        int subColumnCount = 0;
        while (keys.hasNext()) {
            String key = keys.next();
            Cell headerCell = headerRow.createCell(columnCount);
            headerCell.setCellValue(key);

            // todo refactor
            // if exists sub header
            if (jsonObject.get(key) instanceof JSONObject) {
                JSONObject subObject = jsonObject.getJSONObject(key);
                Iterator<String> subKeys = subObject.keys();
                while (subKeys.hasNext()) {
                    String subKey = subKeys.next();
                    Cell subheaderCell = subheaderRow.createCell(subColumnCount);
                    subheaderCell.setCellValue(subKey);
                    subColumnCount++;
                    columnCount++;
                }
                columnCount--;
            }
            columnCount++;
        }
    }

    // todo refactor
    private void createHeader(JSONArray jsonArray, Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        Row subheaderRow = sheet.createRow(1);
        JSONObject jsonObject = jsonArray.getJSONObject(0);
        Iterator<String> keys = jsonObject.keys();
        int columnCount = 0;
        int subColumnCount = 0;
        while (keys.hasNext()) {
            String key = keys.next();
            Cell headerCell = headerRow.createCell(columnCount);
            headerCell.setCellValue(key);

            // todo refactor
            // if exists sub header
            if (jsonObject.get(key) instanceof JSONObject) {
                JSONObject subObject = jsonObject.getJSONObject(key);
                Iterator<String> subKeys = subObject.keys();
                while (subKeys.hasNext()) {
                    String subKey = subKeys.next();
                    Cell subheaderCell = subheaderRow.createCell(subColumnCount);
                    subheaderCell.setCellValue(subKey);
                    subColumnCount++;
                    columnCount++;
                }
                columnCount--;
            }
            columnCount++;
        }
    }
}
