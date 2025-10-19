package Utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

public class ExtractCellValue {

    public static Object extractCellValue(Cell cell) {
        if (cell == null) return null;

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();

            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    yield cell.getDateCellValue();
                } else {
                    double d = cell.getNumericCellValue();
                    // return as Integer if it's a whole number
                    yield (d == Math.floor(d)) ? (int) d : d;
                }
            }

            case BOOLEAN -> cell.getBooleanCellValue();

            case FORMULA -> {
                // evaluate based on cached result
                yield switch (cell.getCachedFormulaResultType()) {
                    case STRING -> cell.getStringCellValue();
                    case NUMERIC -> {
                        if (DateUtil.isCellDateFormatted(cell)) {
                            yield cell.getDateCellValue();
                        } else {
                            double d = cell.getNumericCellValue();
                            yield (d == Math.floor(d)) ? (int) d : d;
                        }
                    }
                    case BOOLEAN -> cell.getBooleanCellValue();
                    default -> null;
                };
            }

            default -> null;
        };
    }
}
