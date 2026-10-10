# Excel 工具类

基于 Java 17 和 Apache POI，支持注解驱动的 XLSX 导出，以及 XLS/XLSX 导入。

## 构建和测试

```bash
mvn test
```

## 运行示例

```bash
mvn exec:java -Dexec.mainClass=com.skynet.common.excel.ExcelDemo
```

示例会在项目根目录生成 `users.xlsx`，再将其导入并打印结果。也可以通过 `-Dexec.args="/path/to/output.xlsx"` 指定输出文件。

## 字段注解

- `@ExcelColumn`：表头名称、列顺序和列宽。
- `@ExcelDateFormat`：日期导入和导出的格式，例如 `pattern = "yyyy-MM-dd HH:mm:ss"`。
- `@PercentFormat`：以 Excel 百分比格式显示数值，字段按比例存储，例如 `0.125` 显示为 `12.50%`。
- `@DictFormat`：使用 `DictService` 将字典值和显示文本双向转换。
- `@BoolFormat`：将布尔值显示为指定文本，并在导入时转换回来。
