package com.skynet.common.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 导出时将数值按 Excel 百分比格式显示。
 * 字段值按比例存储，例如 0.125 显示为 12.50%；字符串也可使用比例值（"0.125"）或带百分号的值（"12.5%"）。
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface PercentFormat {

    /** 百分号后保留的小数位数。 */
    int scale() default 2;
}
