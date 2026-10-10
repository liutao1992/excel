package com.skynet.common.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/** Small file-based helpers used by the runnable demo. */
public final class ExcelTestUtils {

    private ExcelTestUtils() {
    }

    public static <T> void exportTestData(List<T> dataList, String sheetName, Path file) throws IOException {
        try (OutputStream out = Files.newOutputStream(file)) {
            ExcelExportUtil.exportToExcel(dataList, sheetName, out);
        }
    }

    public static <T> List<T> importTestData(Path file, Class<T> clazz) throws IOException {
        try (InputStream in = Files.newInputStream(file)) {
            return ExcelExportUtil.importFromExcel(in, clazz);
        }
    }
}
