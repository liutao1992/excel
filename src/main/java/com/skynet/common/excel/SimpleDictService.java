package com.skynet.common.excel;

import java.util.Map;

public class SimpleDictService implements DictService {

    private static final Map<String, Map<String, String>> DICT_MAP = Map.of(
            "status", Map.of("retrieved", "不在柜", "stored", "在柜")
    );

    @Override
    public String translate(String dictType, String key) {
        return DICT_MAP.getOrDefault(dictType, Map.of()).getOrDefault(key, key);
    }

    @Override
    public String translateReverse(String dictType, String label) {
        return DICT_MAP.getOrDefault(dictType, Map.of()).entrySet().stream()
                .filter(entry -> entry.getValue().equals(label))
                .map(Map.Entry::getKey)
                .findFirst()
                .orElse(label);
    }

}
