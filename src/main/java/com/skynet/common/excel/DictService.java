package com.skynet.common.excel;

public interface DictService {

    /** 正向翻译：code -> label */
    String translate(String dictType, String key);

    /** 反向翻译：label -> code（导入时用） */
    default String translateReverse(String dictType, String label) {
        // 默认回退：不做反向翻译，返回原值
        return label;
    }
}
