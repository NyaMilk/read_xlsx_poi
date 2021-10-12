import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class ExcelReader {
    private Workbook loadWorkbook(String filename) throws IOException, URISyntaxException {
        var extension = filename.substring(filename.lastIndexOf(".") + 1).toLowerCase();
        URL resource = getClass().getClassLoader().getResource(filename);
        var file = new FileInputStream(new File(resource.toURI()));
        switch (extension) {
            case "xls":
                return new HSSFWorkbook(file);
            case "xlsx":
                return new XSSFWorkbook(file);
            default:
                throw new RuntimeException("Неизвестное расширение файла: " + extension);
        }
    }

    private void processRow(HashMap<Integer, List<Object>> data, int rowIndex, Row row) {
        data.put(rowIndex, new ArrayList<>());
        for (var cell : row) {
//            processCell(cell, data.get(rowIndex));
            var value = cell.getStringCellValue();
            var sortedString = Stream.of(value.toLowerCase().split(""))
                    .sorted()
                    .collect(Collectors.joining());
            data.get(rowIndex).add(cell);
            data.get(rowIndex).add(sortedString);
        }
    }

//    private void processCell(Cell cell, List<Object> dataRow) {
//        switch (cell.getCellType()) {
//            case STRING:
//                dataRow.add(cell.getStringCellValue());
//                break;
//            case NUMERIC:
//                if (DateUtil.isCellDateFormatted(cell)) {
//                    dataRow.add(cell.getLocalDateTimeCellValue());
//                } else {
//                    dataRow.add(NumberToTextConverter.toText(cell.getNumericCellValue()));
//                }
//                break;
//            case BOOLEAN:
//                dataRow.add(cell.getBooleanCellValue());
//                break;
//            case FORMULA:
//                dataRow.add(cell.getCellFormula());
//                break;
//            default:
//                dataRow.add(" ");
//        }
//    }

    private HashMap<Integer, List<Object>> processSheet(Sheet sheet) {
        var data = new HashMap<Integer, List<Object>>();
        var iterator = sheet.rowIterator();
        for (var rowIndex = 0; iterator.hasNext(); rowIndex++) {
            var row = iterator.next();
            processRow(data, rowIndex, row);
        }
//        System.out.println("Sheet data:");
//        System.out.println(data);
        return data;
    }

    public HashMap<Integer, List<Object>> read(String filename) throws IOException, URISyntaxException {
        Workbook workbook = loadWorkbook(filename);
        var sheetIterator = workbook.sheetIterator();
        Sheet sheet = sheetIterator.next();
        return processSheet(sheet);
    }
}
