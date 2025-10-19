package DataAccess;

import Utils.ExtractCellValue;
//import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.util.*;
import Model.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRowDAO<T> {
    private final Class<T> type;

    @SuppressWarnings("unchecked")
    public ExcelRowDAO() {
        this.type = (Class<T>) ((ParameterizedType) getClass().getGenericSuperclass()).getActualTypeArguments()[0];
    }

    public ExcelRowDAO(Class<T> type) {
        this.type = type;
    }

    public List<ExcelRow> readFromExcel(String path) {
        List<ExcelRow> result = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(new File(path));
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            if (!rowIterator.hasNext()) return result; // empty sheet protection

            // safely read headers using ExtractCellValue.extractCellValue()
            Row headerRow = rowIterator.next();
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                Object val = ExtractCellValue.extractCellValue(cell);
                headers.add(val == null ? "" : val.toString());
            }

            // safely read all rows
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                ExcelRow excelRow = new ExcelRow();

                for (int i = 0; i < headers.size(); i++) {
                    String fieldName = headers.get(i);
                    Cell cell = row.getCell(i);
                    Object value = ExtractCellValue.extractCellValue(cell);
                    excelRow.set(fieldName, value);
                }

                result.add(excelRow);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        return result;
    }


    public void writeToExcel(List<ExcelRow> rows, String outputPath) {
        if (rows == null || rows.isEmpty()) {
            System.out.println("No data to write.");
            return;
        }

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Export");

            // bold header, grey background
            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short)11);
            headerStyle.setFont(headerFont);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // create header row
            Row headerRow = sheet.createRow(0);

            List<String> fieldNames = rows.get(0).getFieldNames();
            String[] headers = fieldNames.toArray(new String[0]);

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            // fill data rows
            for (int i = 0; i < rows.size(); i++) {
                Row dataRow = sheet.createRow(i + 1);
                ExcelRow excelRow = rows.get(i);

                for (int j = 0; j < headers.length; j++) {
                    Cell cell = dataRow.createCell(j);
                    Object value = excelRow.get(headers[j]);

                    switch (value) {
                        case null -> cell.setBlank();
                        case Number number -> cell.setCellValue(number.doubleValue());
                        case Boolean b -> cell.setCellValue(b);
                        case Date date -> {
                            CellStyle dateStyle = workbook.createCellStyle();
                            CreationHelper createHelper = workbook.getCreationHelper();
                            dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-mm-dd"));
                            cell.setCellStyle(dateStyle);
                            cell.setCellValue(date);
                        }
                        default -> cell.setCellValue(value.toString());
                    }
                }
            }

            // autosize columns
            for (int i = 0; i < headers.length; i++)
                sheet.autoSizeColumn(i);


            // write to file
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                workbook.write(fos);
                System.out.println("Data successfully written to " + outputPath);
            } catch (IOException e) {
                System.err.println("Error writing Excel file: " + e.getMessage());
                e.printStackTrace();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}
