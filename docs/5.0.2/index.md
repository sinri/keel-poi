---
title: Keel POI 5.0.2 用户使用说明
description: 在 Keel / Vert.x 异步模型下读写 Excel（XLS/XLSX）与 CSV 的集成库说明
---

# Keel POI 5.0.2 用户使用说明

[Keel POI](https://github.com/sinri/keel-poi)（
`io.github.sinri:keel-poi`）是在 [Keel](https://github.com/sinri/keel-core) 与 Vert.x
`Future` 异步模型之上，对 Apache POI、excel-streaming-reader 和自研 CSV 读写能力的封装库，用于以统一风格打开或创建工作簿、按行或矩阵读写表数据，并可选处理超大 XLSX 流式读取。

本文档面向 **5.0.2** 版本，适用于托管在 GitHub Pages（例如站点根目录指向 `docs/`）时的版本化说明页。

- 设计与实现审查（工程向）：[5.0.2 审查报告](./risks.md)

---

## 5.0.2 变更摘要

5.0.2 是面向 5.0.x API 的稳定性与依赖维护版本，主要变更如下：

- 自动关闭包装器更稳健：`KeelCsvReader.read(...)`、`KeelCsvWriter.write(...)`、
  `KeelSheets.useSheets(...)` 在用户回调同步抛异常时也会关闭底层资源。
- Excel 行过滤迭代器更符合过滤语义：`getRawRowIterator(...)`、`getMatrixRowIterator(...)`、
  `getTemplatedMatrixRowIterator(...)` 会跳过被 `SheetRowFilter` 排除的行，不再向调用者返回 `null` 行。
- 模板化行按列名读取尾部缺失列时与按索引读取保持一致：列存在但数据行缺少该列时返回空字符串；列名不存在时抛出包含列名的
  `IllegalArgumentException`。
- 矩阵与模板对象加强封装：添加行数据时会保存副本；返回的行列表和列名列表为只读视图，避免调用方意外修改内部状态。
- 补充回归测试：覆盖同步异常自动关闭、行过滤迭代、模板化缺失列、矩阵防御性拷贝、XLSX 流式读取、图片提取等路径。
- 依赖线更新：`keel-core` 升级到 `5.0.3`，`excel-streaming-reader` 升级到 `5.2.0`，测试依赖 `keel-test` 升级到 `5.0.4`。

---

## 环境要求

- JDK 17（与当前工程 toolchain 一致）。
- 运行时需具备 Vert.x Core 提供的 `io.vertx.core.Future` 等 API。本库通过 `api` 依赖传递 `keel-core`，通常只需声明
  `keel-poi`。
- 许可证：GPL-3.0（见仓库与已发布 POM）。

---

## 在项目中引入依赖

Maven：

```xml
<dependency>
    <groupId>io.github.sinri</groupId>
    <artifactId>keel-poi</artifactId>
    <version>5.0.2</version>
</dependency>
```

Gradle（Kotlin DSL）：

```kotlin
implementation("io.github.sinri:keel-poi:5.0.2")
```

> 若还未发布到 Maven Central，可临时使用本地构建产物；发布后以 Maven Central 或项目 README 为准。

---

## 核心概念

| 类型                                             | 作用                                               |
|------------------------------------------------|--------------------------------------------------|
| `KeelSheets`                                   | 工作簿入口：打开已有文件或创建新工作簿，用完后在 `useSheets` 的回调链末尾自动关闭。 |
| `SheetsOpenOptions` / `SheetsCreateOptions`    | 打开与新建时的选项：文件或流、是否 XLSX、公式求值器、大文件流式读取、XLSX 流式写入等。 |
| `KeelSheet`                                    | 单张工作表的读写：矩阵、模板化矩阵、按行遍历、写入单元格、读取图片等。              |
| `KeelSheetMatrix` / `KeelSheetTemplatedMatrix` | 表头 + 数据行的内存模型；模板化矩阵可按列名访问。                       |
| `KeelCsvReader` / `KeelCsvWriter`              | CSV 按行解析与写入，支持分隔符、字符集和 RFC 4180 风格转义。            |

公共 API 多使用 `Future` 表达异步边界；阻塞式方法（如 `readAllRowsToMatrix`）适合在 Worker 线程或简单脚本中使用。

---

## Excel：打开已有工作簿

使用 `KeelSheets.useSheets(SheetsOpenOptions, usage)`。`usage` 返回的 `Future` 完成后，库会关闭工作簿；5.0.2 起，`usage` 在返回
`Future` 前同步抛异常时也会触发关闭。

从文件打开：

```java
import io.github.sinri.keel.integration.poi.excel.*;
import io.github.sinri.keel.integration.poi.excel.entity.KeelSheetMatrix;
import io.vertx.core.Future;

Future<KeelSheetMatrix> future = KeelSheets.useSheets(
        new SheetsOpenOptions()
                .setFile("/path/to/workbook.xlsx")
                .setWithFormulaEvaluator(true),
        sheets -> {
            KeelSheet sheet = sheets.generateReaderForSheet("Sheet1");
            KeelSheetMatrix matrix = sheet.readAllRowsToMatrix();
            return Future.succeededFuture(matrix);
        }
);
```

从输入流打开：

- 若 `setUseXlsx(null)`（默认），库会先把流读入字节数组再尝试 XLSX / XLS 判别，不适合极大文件。
- `setInputStream` 与 `setFile` 互斥：设置其中一个会清空另一个，避免隐式优先级困扰。

超大 XLSX 流式读取（基于 excel-streaming-reader）：

```java
SheetsOpenOptions options = new SheetsOpenOptions()
        .setFile("/path/to/large.xlsx")
        .setHugeXlsxStreamingReaderBuilder(builder -> builder.rowCacheSize(100));

Future<Void> done = KeelSheets.useSheets(options, sheets -> {
    KeelSheet sheet = sheets.generateReaderForSheet("Data");
    Iterator<List<String>> rows = sheet.getRawRowIterator(20, SheetRowFilter.toExcludeEmptyRows());
    while (rows.hasNext()) {
        List<String> row = rows.next();
        // handle row
    }
    return Future.succeededFuture();
});
```

流式模式下可逐行访问当前行单元格，不能像普通 XSSF 一样随机跳行。

---

## Excel：创建新工作簿并保存

`KeelSheets.useSheets(SheetsCreateOptions, usage)` 在结束后同样会关闭工作簿。

```java
SheetsCreateOptions opts = new SheetsCreateOptions()
        .setUseXlsx(true)
        .setUseStreamWriting(true)
        .setWithFormulaEvaluator(false);

Future<Void> done = KeelSheets.useSheets(opts, sheets -> {
    KeelSheet writer = sheets.generateWriterForSheet("Data");

    KeelSheetMatrix matrix = new KeelSheetMatrix();
    matrix.setHeaderRow(List.of("A", "B"));
    matrix.addRow(List.of("1", "2"));

    writer.writeMatrix(matrix);
    sheets.save("/path/to/out.xlsx");
    return Future.succeededFuture();
});
```

`setUseStreamWriting(true)` 仅对 XLSX 生效；使用 HSSF（`setUseXlsx(false)`）时不会启用 SXSSF。

---

## Excel：读取为矩阵或模板化矩阵

默认约定（`readAllRowsToMatrix()` / `readAllRowsToTemplatedMatrix()`）：

1. 第一行作为表头。使用带 `headerRowIndex` 的重载可指定表头行号。
2. 若 `maxColumns <= 0`，在表头行上用首个非空列范围自动推断列数。
3. 默认排除全空单元格行，见 `SheetRowFilter`。

重载示例：

```java
// headerRowIndex：表头所在行号（0-based）
// maxColumns：> 0 为固定列数；<= 0 为自动检测
// sheetRowFilter：可为 null，表示不过滤
KeelSheetMatrix matrix = sheet.readAllRowsToMatrix(
        0,
        0,
        SheetRowFilter.toExcludeEmptyRows()
);
```

异步批量读取需传入 `Keel` 实例，按批处理 POI `Row`：

- `readAllRowsToMatrixAsync(keel, ...)`
- `readAllRowsToTemplatedMatrixAsync(keel, ...)`

矩阵转模板化矩阵：

```java
KeelSheetTemplatedMatrix templated = matrix.transformToTemplatedMatrix();
```

表头为空时会抛异常。表头中存在重复列名时，`KeelSheetMatrixRowTemplate.create(...)` 会抛出 `IllegalArgumentException`。

---

## 矩阵与模板化行的 5.0.2 行为

5.0.2 对矩阵内部状态做了更严格的保护：

- `KeelSheetMatrix.addRow(...)`、`addRows(...)` 会复制传入的行列表。
- `KeelSheetTemplatedMatrix.addRawRow(...)` 和 `KeelSheetMatrixTemplatedRow.create(...)` 会复制传入的行列表。
- `getRawRow(...)`、`getRawRowList()`、`getRawRows()`、`getRawRow()`、`getColumnNames()` 返回只读视图。

这意味着调用方不能再通过 getter 返回的列表直接修改矩阵内容。需要变更数据时，应构造新的行或矩阵后再写入。

模板化行读取规则：

```java
KeelSheetMatrixRowTemplate template = KeelSheetMatrixRowTemplate.create(List.of("A", "B", "C"));
KeelSheetMatrixTemplatedRow row = KeelSheetMatrixTemplatedRow.create(template, List.of("a", "b"));

row.getColumnValue(2);   // ""
row.getColumnValue("C"); // ""
row.getColumnValue("D"); // IllegalArgumentException
```

`KeelSheetMatrixRow` 的 `readValueToInteger` / `readValueToLong` 等数值 API 在无法解析或越界时返回 `null`。

---

## Excel：写入矩阵与原始行

- `writeMatrix(KeelSheetMatrix)`：有表头则第 0 行写表头，数据从第 1 行起；无表头则只写数据行。
- `writeTemplatedMatrix(...)` / 对应的 `*Async`：写入模板化矩阵。
- `writeAllRows(List<List<String>> rowData, sinceRowIndex, sinceCellIndex)`：低开销直接按行列下标写字符串。

---

## Excel：公式与图片

- 公式：在 `SheetsOpenOptions` / `SheetsCreateOptions` 上设置 `setWithFormulaEvaluator(true)`，并在
  `generateReaderForSheet` 时使用默认或 `parseFormulaCellToValue == true`，可将公式单元格按求值结果读出。
- 图片：`KeelSheet.getPictures()` 返回
  `KeelPictureInSheet` 列表，用于从绘图层提取内嵌图片。5.0.2 回归测试覆盖了 XLSX 图片元数据与数据副本读取路径。

---

## 行过滤：`SheetRowFilter`

内置排除全空行：

```java
SheetRowFilter filter = SheetRowFilter.toExcludeEmptyRows();
```

- 推荐使用 `shouldExcludeRow(rawRow)`，返回 `true` 表示丢弃该行。
- 旧名 `shouldThrowThisRawRow` / `toThrowEmptyRows` 在 5.0.1 起标记为 `@Deprecated`，语义不变，建议迁移到新命名。
- 5.0.2 起，迭代器会跳过被过滤行；`next()` 返回实际保留的行，不再返回 `null`。

---

## CSV：读取

推荐使用静态方法 `KeelCsvReader.read(InputStream, Charset, separator, readFunc)`，由库在结束时关闭读取器。5.0.2 起，
`readFunc` 同步抛异常时也会关闭读取器。

```java
KeelCsvReader.read(
        inputStream,
        StandardCharsets.UTF_8,
        ",",
        reader -> {
            try {
                CsvRow row;
                while ((row = reader.next()) != null) {
                    row.stream().forEach(cell -> {
                        String value = cell.getString();
                        // handle value
                    });
                }
                return Future.succeededFuture();
            } catch (IOException e) {
                return Future.failedFuture(e);
            }
        }
);
```

解析支持带引号字段、分隔符、换行和双引号转义等常见 CSV 变体。

---

## CSV：写入

`KeelCsvWriter.write(...)` 会在写入回调完成后关闭输出流。5.0.2 起，写入回调同步抛异常时也会关闭输出流。

```java
KeelCsvWriter.write(outputStream, writer -> {
    try {
        writer.blockWriteRow(List.of("name", "age"));
        writer.writeCell("Alice");
        writer.writeCell("18");
        writer.writeRowEnding();
        return Future.succeededFuture();
    } catch (IOException e) {
        return Future.failedFuture(e);
    }
});
```

- `write(OutputStream, writeFunc)` 使用默认 UTF-8 和逗号分隔。
- `write(OutputStream, separator, Charset, writeFunc)` 可指定分隔符和字符集。
- 逻辑行由 `writeCell` 多次调用 + `writeRowEnding()` 组成，或使用 `blockWriteRow(List)` 一次性写一行。
- 行分隔符为 CRLF（`\r\n`），贴近 RFC 4180。

---

## 兼容性提示

5.0.2 不引入新的公开入口类，但以下行为比 5.0.1 更严格：

- 通过 getter 返回的列表直接修改矩阵或模板内部状态的用法不再可行，会抛出 `UnsupportedOperationException`。
- 行过滤迭代器不再把被过滤行表现为 `null`。如果既有代码依赖 `next() == null` 判断过滤行，需要改为只处理迭代器返回的实际行。
- `getColumnValue(String)` 对不存在的列名明确抛出 `IllegalArgumentException`；对存在但数据行尾部缺失的列返回空字符串。

---

## 验证

5.0.2 分支当前验证结果：

```text
./gradlew test
```

测试通过。详见 [5.0.2 审查报告](./risks.md) 中的处理状态与覆盖说明。
