package com.skynet.common.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 导入和导出时使用的日期格式。
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelDateFormat {

    /** 日期格式，例如 yyyy-MM-dd HH:mm:ss。 */
    String pattern();
}
