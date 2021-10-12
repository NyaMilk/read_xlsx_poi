import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;

public class ExcelWriter {
//    private void createHeader(XSSFWorkbook workbook, Sheet sheet) {
//        var header = sheet.createRow(0);
//
//        var headerStyle = workbook.createCellStyle();
//        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
//        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//
//        var font = workbook.createFont();
//        font.setFontName("Arial");
//        font.setFontHeightInPoints((short) 14);
//        font.setBold(true);
//        headerStyle.setFont(font);
//
//        var headerCell = header.createCell(0);
//        headerCell.setCellValue("Word");
//        headerCell.setCellStyle(headerStyle);
//    }

    private void createCells(XSSFWorkbook workbook, Sheet sheet, HashMap<Integer, List<Object>> data) {
        var style = workbook.createCellStyle();
        style.setWrapText(true);
        sheet.setDefaultColumnStyle(1, style);

        for (var i = 0; i < data.size(); i++) {
            var value = data.get(i);
//            var row = sheet.createRow(i + 1);
            var row = sheet.createRow(i);

            var cell = row.createCell(0);
            cell.setCellValue(String.valueOf(value.get(0)));

            cell = row.createCell(1);
            cell.setCellValue(String.valueOf(value.get(1)));
        }
    }

    public void write(String filename, HashMap<Integer, List<Object>> data) throws IOException {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Result");
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 4000);
//        createHeader(workbook, sheet);
        createCells(workbook, sheet, data);
        try (var outputStream = new FileOutputStream(filename)) {
            workbook.write(outputStream);
        }
        workbook.close();
    }
}
