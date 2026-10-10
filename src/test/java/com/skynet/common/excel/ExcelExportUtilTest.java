package com.skynet.common.excel;

import com.skynet.common.excel.annotation.BoolFormat;
import com.skynet.common.excel.annotation.DictFormat;
import com.skynet.common.excel.annotation.ExcelColumn;
import com.skynet.common.excel.annotation.ExcelDateFormat;
import com.skynet.common.excel.annotation.PercentFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelExportUtilTest {

    @BeforeEach
    void configureDictionary() {
        ExcelExportUtil.setDictService(new SimpleDictService());
    }

    @AfterEach
    void clearDictionary() {
        ExcelExportUtil.setDictService(null);
    }

    @Test
    void exportsAndImportsAnnotatedValues() throws Exception {
        LocalDateTime registerTime = LocalDateTime.of(2025, 1, 2, 3, 4, 5);
        SampleRow source = new SampleRow("张三", "retrieved", registerTime,
                new BigDecimal("0.1234"), true);
        byte[] bytes = export(List.of(source));

        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("姓名", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("状态", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals("登记时间", sheet.getRow(0).getCell(2).getStringCellValue());
            assertEquals("完成率", sheet.getRow(0).getCell(3).getStringCellValue());
            assertEquals("有效", sheet.getRow(0).getCell(4).getStringCellValue());

            assertEquals("不在柜", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("2025-01-02 03:04:05", sheet.getRow(1).getCell(2).getStringCellValue());
            assertEquals(CellType.NUMERIC, sheet.getRow(1).getCell(3).getCellType());
            assertEquals(0.1234, sheet.getRow(1).getCell(3).getNumericCellValue(), 0.000001);
            assertEquals("0.00%", sheet.getRow(1).getCell(3).getCellStyle().getDataFormatString());
            assertEquals("是", sheet.getRow(1).getCell(4).getStringCellValue());
        }

        SampleRow imported = ExcelExportUtil.importFromExcel(
                new ByteArrayInputStream(bytes), SampleRow.class).get(0);
        assertEquals(source.name, imported.name);
        assertEquals(source.status, imported.status);
        assertEquals(source.registerTime, imported.registerTime);
        assertEquals(0, source.completionRate.compareTo(imported.completionRate));
        assertEquals(source.active, imported.active);
    }

    @Test
    void importsXlsAndMapsColumnsByHeaderInsteadOfPosition() throws Exception {
        byte[] bytes;
        try (Workbook workbook = new HSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            var sheet = workbook.createSheet("input");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue("完成率");
            header.createCell(1).setCellValue("姓名");
            var row = sheet.createRow(1);
            row.createCell(0).setCellValue(0.75);
            row.createCell(1).setCellValue("李四");
            workbook.write(out);
            bytes = out.toByteArray();
        }

        SampleRow imported = ExcelExportUtil.importFromExcel(
                new ByteArrayInputStream(bytes), SampleRow.class).get(0);
        assertEquals("李四", imported.name);
        assertEquals(0, new BigDecimal("0.75").compareTo(imported.completionRate));
        assertNull(imported.status);
    }

    @Test
    void exportsNullValuesAsBlankCells() throws Exception {
        byte[] bytes = export(List.of(new SampleRow("张三", null, null, null, null)));

        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            var row = workbook.getSheetAt(0).getRow(1);
            assertEquals("", row.getCell(1).getStringCellValue());
            assertEquals("", row.getCell(2).getStringCellValue());
            assertEquals("", row.getCell(3).getStringCellValue());
            assertEquals("", row.getCell(4).getStringCellValue());
        }
    }

    @Test
    void rejectsEmptyExportsAndInvalidPercentStrings() {
        assertThrows(IllegalArgumentException.class,
                () -> ExcelExportUtil.exportToExcel(List.of(), "empty", new ByteArrayOutputStream()));
        assertThrows(RuntimeException.class,
                () -> ExcelExportUtil.exportToExcel(List.of(new InvalidPercentRow("not-a-number")),
                        "invalid", new ByteArrayOutputStream()));
    }

    @Test
    void rejectsUnknownBooleanLabelsDuringImport() throws Exception {
        byte[] bytes;
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            var sheet = workbook.createSheet("input");
            sheet.createRow(0).createCell(0).setCellValue("有效");
            sheet.createRow(1).createCell(0).setCellValue("未知");
            workbook.write(out);
            bytes = out.toByteArray();
        }

        assertThrows(RuntimeException.class,
                () -> ExcelExportUtil.importFromExcel(new ByteArrayInputStream(bytes), SampleRow.class));
    }

    @Test
    void roundTripsDictionaryAndBooleanValues() throws Exception {
        byte[] bytes = export(List.of(new SampleRow("张三", "stored", null,
                new BigDecimal("0.5"), false)));
        SampleRow imported = ExcelExportUtil.importFromExcel(
                new ByteArrayInputStream(bytes), SampleRow.class).get(0);
        assertFalse(imported.active);
        assertEquals("stored", imported.status);
        assertTrue(imported.active == Boolean.FALSE);
    }

    @Test
    void runnableDemoCreatesAndReadsAnExcelFile(@TempDir Path tempDir) throws Exception {
        Path output = tempDir.resolve("demo.xlsx");

        ExcelDemo.main(new String[]{output.toString()});

        assertTrue(Files.size(output) > 0);
    }

    private static byte[] export(List<SampleRow> rows) throws Exception {
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            ExcelExportUtil.exportToExcel(rows, "测试数据", out);
            return out.toByteArray();
        }
    }

    public static class SampleRow {
        @ExcelColumn(value = "姓名", order = 0)
        private String name;

        @ExcelColumn(value = "状态", order = 1)
        @DictFormat(dictType = "status")
        private String status;

        @ExcelColumn(value = "登记时间", order = 2)
        @ExcelDateFormat(pattern = "yyyy-MM-dd HH:mm:ss")
        private LocalDateTime registerTime;

        @ExcelColumn(value = "完成率", order = 3)
        @PercentFormat(scale = 2)
        private BigDecimal completionRate;

        @ExcelColumn(value = "有效", order = 4)
        @BoolFormat(trueText = "是", falseText = "否")
        private Boolean active;

        public SampleRow() {
        }

        SampleRow(String name, String status, LocalDateTime registerTime,
                  BigDecimal completionRate, Boolean active) {
            this.name = name;
            this.status = status;
            this.registerTime = registerTime;
            this.completionRate = completionRate;
            this.active = active;
        }
    }

    public static class InvalidPercentRow {
        @ExcelColumn(value = "百分比", order = 0)
        @PercentFormat(scale = 2)
        private String percentage;

        public InvalidPercentRow() {
        }

        InvalidPercentRow(String percentage) {
            this.percentage = percentage;
        }
    }
}
