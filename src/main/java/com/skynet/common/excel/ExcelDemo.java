package com.skynet.common.excel;

import com.skynet.common.excel.annotation.BoolFormat;
import com.skynet.common.excel.annotation.DictFormat;
import com.skynet.common.excel.annotation.ExcelColumn;
import com.skynet.common.excel.annotation.ExcelDateFormat;
import com.skynet.common.excel.annotation.PercentFormat;

import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;

/** Runnable example: exports annotated values, then imports them back. */
public final class ExcelDemo {

    private ExcelDemo() {
    }

    public static void main(String[] args) throws IOException {
        Path file = args.length > 0 ? Path.of(args[0]) : Path.of("users.xlsx");
        Path parent = file.toAbsolutePath().getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        ExcelExportUtil.setDictService(new SimpleDictService());
        List<UserRow> users = List.of(
                new UserRow("张三", "retrieved", LocalDateTime.now(), new BigDecimal("0.8234"), true),
                new UserRow("李四", "stored", LocalDateTime.now().minusDays(1), new BigDecimal("0.6750"), false)
        );

        ExcelTestUtils.exportTestData(users, "用户数据", file);
        List<UserRow> imported = ExcelTestUtils.importTestData(file, UserRow.class);

        System.out.println("Excel 文件已生成: " + file.toAbsolutePath());
        System.out.println("导入记录数: " + imported.size());
        imported.forEach(System.out::println);
    }

    public static class UserRow {
        @ExcelColumn(value = "姓名", order = 0, width = 20)
        private String name;

        @ExcelColumn(value = "在柜状态", order = 1, width = 16)
        @DictFormat(dictType = "status")
        private String status;

        @ExcelColumn(value = "登记时间", order = 2, width = 24)
        @ExcelDateFormat(pattern = "yyyy-MM-dd HH:mm:ss")
        private LocalDateTime registerTime;

        @ExcelColumn(value = "完成率", order = 3, width = 14)
        @PercentFormat(scale = 2)
        private BigDecimal completionRate;

        @ExcelColumn(value = "是否有效", order = 4, width = 14)
        @BoolFormat(trueText = "是", falseText = "否")
        private Boolean active;

        public UserRow() {
        }

        public UserRow(String name, String status, LocalDateTime registerTime,
                       BigDecimal completionRate, Boolean active) {
            this.name = name;
            this.status = status;
            this.registerTime = registerTime;
            this.completionRate = completionRate;
            this.active = active;
        }

        @Override
        public String toString() {
            return "UserRow{" +
                    "name='" + name + '\'' +
                    ", status='" + status + '\'' +
                    ", registerTime=" + registerTime +
                    ", completionRate=" + completionRate +
                    ", active=" + active +
                    '}';
        }
    }
}
