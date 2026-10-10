package com.skynet.common.excel;

import com.skynet.common.excel.annotation.BoolFormat;
import com.skynet.common.excel.annotation.DictFormat;
import com.skynet.common.excel.annotation.ExcelColumn;
import com.skynet.common.excel.annotation.ExcelDateFormat;
import com.skynet.common.excel.annotation.PercentFormat;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.ParseException;
import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * 基于字段注解导入和导出 Excel。
 * 导出使用 XLSX；导入支持 XLS 和 XLSX。
 */
public final class ExcelExportUtil {

    private static volatile DictService dictService;

    private ExcelExportUtil() {
    }

    public static void setDictService(DictService service) {
        dictService = service;
    }

    /**
     * 将对象列表导出为 XLSX。表头和值只处理带有 {@link ExcelColumn} 的字段。
     */
    public static <T> void exportToExcel(List<T> dataList, String sheetName, OutputStream outputStream) {
        if (dataList == null || dataList.isEmpty()) {
            throw new IllegalArgumentException("数据列表不能为空");
        }
        Objects.requireNonNull(outputStream, "输出流不能为空");
        if (sheetName == null || sheetName.isBlank()) {
            throw new IllegalArgumentException("Sheet 名称不能为空");
        }

        try (Workbook workbook = new XSSFWorkbook()) {
            Class<?> rowType = Objects.requireNonNull(dataList.get(0), "数据行不能为空").getClass();
            List<Field> fields = getExcelFields(rowType);
            validateUniqueHeaders(fields);
            if (fields.isEmpty()) {
                throw new IllegalArgumentException("类型中没有标注 @ExcelColumn 的字段: " + rowType.getName());
            }

            Sheet sheet = workbook.createSheet(sheetName);
            Map<Integer, CellStyle> percentStyles = new HashMap<>();

            Row headerRow = sheet.createRow(0);
            for (int columnIndex = 0; columnIndex < fields.size(); columnIndex++) {
                Field field = fields.get(columnIndex);
                ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                headerRow.createCell(columnIndex).setCellValue(column.value());
                sheet.setColumnWidth(columnIndex, column.width() * 256);
            }

            for (int rowIndex = 0; rowIndex < dataList.size(); rowIndex++) {
                T item = Objects.requireNonNull(dataList.get(rowIndex), "第 " + (rowIndex + 1) + " 行数据不能为空");
                if (!rowType.isInstance(item)) {
                    throw new IllegalArgumentException("数据列表中的对象类型不一致");
                }

                Row row = sheet.createRow(rowIndex + 1);
                for (int columnIndex = 0; columnIndex < fields.size(); columnIndex++) {
                    Field field = fields.get(columnIndex);
                    field.setAccessible(true);
                    Object rawValue = field.get(item);
                    Cell cell = row.createCell(columnIndex);

                    PercentFormat percentFormat = field.getAnnotation(PercentFormat.class);
                    if (percentFormat != null) {
                        writePercentCell(cell, rawValue, field, percentFormat, workbook, percentStyles);
                        continue;
                    }

                    cell.setCellValue(formatValueForExport(rawValue, field));
                }
            }

            workbook.write(outputStream);
        } catch (Exception e) {
            throw new RuntimeException("导出 Excel 失败", e);
        }
    }

    private static void writePercentCell(Cell cell, Object rawValue, Field field,
                                         PercentFormat format, Workbook workbook,
                                         Map<Integer, CellStyle> styles) {
        cell.setCellStyle(getPercentStyle(workbook, format.scale(), styles));
        if (rawValue == null) {
            cell.setCellValue("");
            return;
        }
        if (rawValue instanceof Number) {
            double numericValue = ((Number) rawValue).doubleValue();
            if (!Double.isFinite(numericValue)) {
                throw new IllegalArgumentException("百分比字段值超出 Excel 数值范围: " + field.getName());
            }
            cell.setCellValue(numericValue);
            return;
        }
        if (!(rawValue instanceof String)) {
            throw new IllegalArgumentException("百分比字段必须是数值类型或数字字符串: " + field.getName());
        }

        String textValue = ((String) rawValue).trim();
        if (textValue.isEmpty()) {
            cell.setCellValue("");
            return;
        }
        boolean percentLiteral = textValue.endsWith("%");
        String numericText = percentLiteral
                ? textValue.substring(0, textValue.length() - 1).trim()
                : textValue;
        BigDecimal decimalValue;
        try {
            decimalValue = new BigDecimal(numericText);
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException("百分比字段必须是有效数字或百分比字符串: " + field.getName(), e);
        }
        if (percentLiteral) {
            decimalValue = decimalValue.movePointLeft(2);
        }
        double numericValue = decimalValue.doubleValue();
        if (!Double.isFinite(numericValue)) {
            throw new IllegalArgumentException("百分比字段值超出 Excel 数值范围: " + field.getName());
        }
        cell.setCellValue(numericValue);
    }

    private static CellStyle getPercentStyle(Workbook workbook, int scale, Map<Integer, CellStyle> styles) {
        if (scale < 0 || scale > 30) {
            throw new IllegalArgumentException("百分比小数位数必须在 0 到 30 之间: " + scale);
        }
        return styles.computeIfAbsent(scale, digits -> {
            CellStyle style = workbook.createCellStyle();
            String numberFormat = digits == 0 ? "0%" : "0." + "0".repeat(digits) + "%";
            style.setDataFormat(workbook.createDataFormat().getFormat(numberFormat));
            return style;
        });
    }

    private static String formatValueForExport(Object rawValue, Field field) {
        if (rawValue == null) {
            return "";
        }

        BoolFormat boolFormat = field.getAnnotation(BoolFormat.class);
        if (boolFormat != null) {
            if (!(rawValue instanceof Boolean)) {
                throw new IllegalArgumentException("@BoolFormat 字段必须是 Boolean 类型: " + field.getName());
            }
            return (Boolean) rawValue ? boolFormat.trueText() : boolFormat.falseText();
        }

        DictFormat dictFormat = field.getAnnotation(DictFormat.class);
        if (dictFormat != null && dictService != null) {
            String translated = dictService.translate(dictFormat.dictType(), rawValue.toString());
            return translated == null ? "" : translated;
        }

        ExcelDateFormat dateFormat = field.getAnnotation(ExcelDateFormat.class);
        String pattern = dateFormat == null ? "" : dateFormat.pattern();
        if (pattern.isBlank()) {
            return rawValue.toString();
        }
        if (rawValue instanceof Date) {
            return new SimpleDateFormat(pattern).format((Date) rawValue);
        }
        if (rawValue instanceof LocalDateTime) {
            return ((LocalDateTime) rawValue).format(DateTimeFormatter.ofPattern(pattern));
        }
        if (rawValue instanceof LocalDate) {
            return ((LocalDate) rawValue).format(DateTimeFormatter.ofPattern(pattern));
        }
        if (rawValue instanceof LocalTime) {
            return ((LocalTime) rawValue).format(DateTimeFormatter.ofPattern(pattern));
        }
        return rawValue.toString();
    }

    /**
     * 从 Excel 导入对象列表。Excel 表头与 {@link ExcelColumn#value()} 匹配，列顺序可以不同。
     */
    public static <T> List<T> importFromExcel(InputStream in, Class<T> clazz,
                                              int sheetIndex, int headerRowIndex) {
        Objects.requireNonNull(in, "输入流不能为空");
        Objects.requireNonNull(clazz, "目标类型不能为空");
        if (sheetIndex < 0 || headerRowIndex < 0) {
            throw new IllegalArgumentException("Sheet 和表头行索引不能小于 0");
        }

        try {
            // 防止压缩炸弹。应用可按文件来源和大小调整此阈值。
            ZipSecureFile.setMinInflateRatio(0.001);

            try (Workbook workbook = WorkbookFactory.create(in)) {
                if (sheetIndex >= workbook.getNumberOfSheets()) {
                    throw new IllegalArgumentException("Sheet 序号超出范围: " + sheetIndex);
                }
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                Row headerRow = sheet.getRow(headerRowIndex);
                if (headerRow == null) {
                    throw new IllegalArgumentException("未找到表头行（index=" + headerRowIndex + "）");
                }

                List<Field> fields = getExcelFields(clazz);
                validateUniqueHeaders(fields);
                Map<String, Field> headerToField = new HashMap<>();
                for (Field field : fields) {
                    ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                    headerToField.put(column.value().trim(), field);
                }

                Map<Integer, Field> columnToField = new HashMap<>();
                for (int columnIndex = 0; columnIndex < headerRow.getLastCellNum(); columnIndex++) {
                    String header = getCellAsString(headerRow.getCell(columnIndex)).trim();
                    Field field = headerToField.get(header);
                    if (field != null) {
                        field.setAccessible(true);
                        columnToField.put(columnIndex, field);
                    }
                }
                if (columnToField.isEmpty()) {
                    throw new IllegalArgumentException("Excel 表头中没有匹配到 @ExcelColumn 字段");
                }

                Constructor<T> constructor = clazz.getDeclaredConstructor();
                constructor.setAccessible(true);
                List<T> result = new ArrayList<>();
                for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (isRowEmpty(row)) {
                        continue;
                    }

                    T instance = constructor.newInstance();
                    for (Map.Entry<Integer, Field> entry : columnToField.entrySet()) {
                        int columnIndex = entry.getKey();
                        Field field = entry.getValue();
                        Cell cell = row.getCell(columnIndex);
                        String value = extractCellString(cell);

                        BoolFormat boolFormat = field.getAnnotation(BoolFormat.class);
                        if (boolFormat != null && !value.isBlank()) {
                            value = normalizeBooleanText(value, boolFormat);
                        }

                        DictFormat dictFormat = field.getAnnotation(DictFormat.class);
                        if (dictFormat != null && dictService != null) {
                            value = dictService.translateReverse(dictFormat.dictType(), value);
                        }

                        ExcelDateFormat dateFormat = field.getAnnotation(ExcelDateFormat.class);
                        String pattern = dateFormat == null ? "" : dateFormat.pattern();
                        Object converted = convertToFieldType(value, cell, field.getType(), pattern);
                        field.set(instance, converted);
                    }
                    result.add(instance);
                }
                return result;
            }
        } catch (Exception e) {
            throw new RuntimeException("导入 Excel 失败", e);
        }
    }

    /** 默认读取第一个 Sheet，第一行为表头。 */
    public static <T> List<T> importFromExcel(InputStream in, Class<T> clazz) {
        return importFromExcel(in, clazz, 0, 0);
    }

    private static List<Field> getExcelFields(Class<?> clazz) {
        return Arrays.stream(clazz.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                .sorted(Comparator.comparingInt(field -> field.getAnnotation(ExcelColumn.class).order()))
                .toList();
    }

    private static void validateUniqueHeaders(List<Field> fields) {
        Set<String> headers = new HashSet<>();
        for (Field field : fields) {
            String header = field.getAnnotation(ExcelColumn.class).value().trim();
            if (!headers.add(header)) {
                throw new IllegalArgumentException("Excel 表头名称重复: " + header);
            }
        }
    }

    private static boolean isRowEmpty(Row row) {
        if (row == null) {
            return true;
        }
        short firstCellNum = row.getFirstCellNum();
        if (firstCellNum < 0) {
            return true;
        }
        for (int columnIndex = firstCellNum; columnIndex < row.getLastCellNum(); columnIndex++) {
            Cell cell = row.getCell(columnIndex);
            if (cell != null && cell.getCellType() != CellType.BLANK
                    && !getCellAsString(cell).isBlank()) {
                return false;
            }
        }
        return true;
    }

    private static String getCellAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(cell.getDateCellValue());
                }
                return NumberToTextConverter.toText(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return getFormulaCellAsString(cell);
            default:
                return "";
        }
    }

    private static String getFormulaCellAsString(Cell cell) {
        CellType cachedType = cell.getCachedFormulaResultType();
        if (cachedType == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(cell.getDateCellValue());
            }
            return NumberToTextConverter.toText(cell.getNumericCellValue());
        }
        if (cachedType == CellType.STRING) {
            return cell.getStringCellValue();
        }
        if (cachedType == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        }
        return "";
    }

    private static String extractCellString(Cell cell) {
        return getCellAsString(cell).trim();
    }

    private static String normalizeBooleanText(String text, BoolFormat format) {
        if (text.equals(format.trueText())) {
            return "true";
        }
        if (text.equals(format.falseText())) {
            return "false";
        }
        return text;
    }

    private static Object convertToFieldType(String text, Cell cell, Class<?> targetType,
                                             String datePattern) throws ParseException {
        String value = text == null ? "" : text.trim();
        if (value.isEmpty()) {
            if (!targetType.isPrimitive()) {
                return null;
            }
            if (targetType == boolean.class) return false;
            if (targetType == char.class) return '\0';
            if (targetType == byte.class) return (byte) 0;
            if (targetType == short.class) return (short) 0;
            if (targetType == int.class) return 0;
            if (targetType == long.class) return 0L;
            if (targetType == float.class) return 0F;
            if (targetType == double.class) return 0D;
        }

        if (targetType == Date.class) {
            if (isExcelDateCell(cell)) {
                return cell.getDateCellValue();
            }
            Date parsed = datePattern.isEmpty() ? tryParseDateCommon(value) : parseDate(value, datePattern);
            if (parsed == null) {
                throw new ParseException("无法解析日期: " + value, 0);
            }
            return parsed;
        }

        if (targetType == LocalDateTime.class) {
            if (isExcelDateCell(cell)) {
                return LocalDateTime.ofInstant(cell.getDateCellValue().toInstant(), ZoneId.systemDefault());
            }
            LocalDateTime parsed = datePattern.isEmpty()
                    ? tryParseLocalDateTimeCommon(value)
                    : LocalDateTime.parse(value, DateTimeFormatter.ofPattern(datePattern));
            if (parsed == null) {
                throw new ParseException("无法解析日期时间: " + value, 0);
            }
            return parsed;
        }

        if (targetType == LocalDate.class) {
            if (isExcelDateCell(cell)) {
                return cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            }
            LocalDate parsed = datePattern.isEmpty()
                    ? tryParseLocalDateCommon(value)
                    : LocalDate.parse(value, DateTimeFormatter.ofPattern(datePattern));
            if (parsed == null) {
                throw new ParseException("无法解析日期: " + value, 0);
            }
            return parsed;
        }

        if (targetType == LocalTime.class) {
            if (isExcelDateCell(cell)) {
                return cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalTime();
            }
            LocalTime parsed = datePattern.isEmpty()
                    ? tryParseLocalTimeCommon(value)
                    : LocalTime.parse(value, DateTimeFormatter.ofPattern(datePattern));
            if (parsed == null) {
                throw new ParseException("无法解析时间: " + value, 0);
            }
            return parsed;
        }

        if (targetType == String.class || targetType == Object.class) return value;
        if (targetType == Byte.class || targetType == byte.class) return Byte.valueOf(value);
        if (targetType == Short.class || targetType == short.class) return Short.valueOf(value);
        if (targetType == Integer.class || targetType == int.class) return Integer.valueOf(value);
        if (targetType == Long.class || targetType == long.class) return Long.valueOf(value);
        if (targetType == Float.class || targetType == float.class) return Float.valueOf(value);
        if (targetType == Double.class || targetType == double.class) return Double.valueOf(value);
        if (targetType == BigInteger.class) return new BigInteger(value);
        if (targetType == BigDecimal.class) return new BigDecimal(value);
        if (targetType == Boolean.class || targetType == boolean.class) return parseBoolean(value);
        if (targetType == Character.class || targetType == char.class) {
            if (value.length() != 1) {
                throw new IllegalArgumentException("字符字段必须恰好包含一个字符: " + value);
            }
            return value.charAt(0);
        }
        if (targetType.isEnum()) {
            @SuppressWarnings({"unchecked", "rawtypes"})
            Object enumValue = Enum.valueOf((Class<? extends Enum>) targetType, value);
            return enumValue;
        }
        throw new IllegalArgumentException("不支持导入的字段类型: " + targetType.getName());
    }

    private static boolean isExcelDateCell(Cell cell) {
        return cell != null && cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell);
    }

    private static boolean parseBoolean(String value) {
        String normalized = value.toLowerCase(Locale.ROOT);
        if (Arrays.asList("true", "1", "yes", "y", "是").contains(normalized)) return true;
        if (Arrays.asList("false", "0", "no", "n", "否").contains(normalized)) return false;
        throw new IllegalArgumentException("无法识别布尔值: " + value);
    }

    private static Date parseDate(String value, String pattern) throws ParseException {
        SimpleDateFormat formatter = new SimpleDateFormat(pattern);
        formatter.setLenient(false);
        ParsePosition position = new ParsePosition(0);
        Date parsed = formatter.parse(value, position);
        if (parsed == null || position.getIndex() != value.length()) {
            throw new ParseException("日期不符合格式 " + pattern + ": " + value,
                    Math.max(position.getErrorIndex(), position.getIndex()));
        }
        return parsed;
    }

    private static Date tryParseDateCommon(String value) {
        for (String pattern : List.of("yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd", "yyyy/MM/dd HH:mm:ss", "yyyy/MM/dd")) {
            try {
                return parseDate(value, pattern);
            } catch (ParseException ignored) {
                // 尝试下一种常见格式。
            }
        }
        return null;
    }

    private static LocalDateTime tryParseLocalDateTimeCommon(String value) {
        for (String pattern : List.of("yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss", "yyyy-MM-dd'T'HH:mm:ss")) {
            try {
                return LocalDateTime.parse(value, DateTimeFormatter.ofPattern(pattern));
            } catch (RuntimeException ignored) {
                // 尝试下一种常见格式。
            }
        }
        LocalDate date = tryParseLocalDateCommon(value);
        return date == null ? null : date.atStartOfDay();
    }

    private static LocalDate tryParseLocalDateCommon(String value) {
        for (String pattern : List.of("yyyy-MM-dd", "yyyy/MM/dd")) {
            try {
                return LocalDate.parse(value, DateTimeFormatter.ofPattern(pattern));
            } catch (RuntimeException ignored) {
                // 尝试下一种常见格式。
            }
        }
        return null;
    }

    private static LocalTime tryParseLocalTimeCommon(String value) {
        for (String pattern : List.of("HH:mm:ss", "HH:mm")) {
            try {
                return LocalTime.parse(value, DateTimeFormatter.ofPattern(pattern));
            } catch (RuntimeException ignored) {
                // 尝试下一种常见格式。
            }
        }
        return null;
    }
}
