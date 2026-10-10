package com.skynet.common.excel.annotation;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {
    String value();         // Excel 列名
    int order();            // 排序
    int width() default 20; // 列宽，单位为字符（*256）
}
