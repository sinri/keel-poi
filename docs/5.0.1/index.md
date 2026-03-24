---
title: Keel POI 5.0.1 用户使用说明
description: 在 Keel / Vert.x 异步模型下读写 Excel（XLS/XLSX）与 CSV 的集成库说明
---

# Keel POI 5.0.1 用户使用说明

[Keel POI](https://github.com/sinri/keel-poi)（`io.github.sinri:keel-poi`）是在 [**Keel
**](https://github.com/sinri/keel-core) 与 [**Vert.x**](https://vertx.io/) 的 `Future` 异步模型之上，对 **Apache POI
** 与自研 **CSV** 读写的封装库，用于以更统一的风格打开/创建工作簿、按行或按矩阵读写表数据，并可选超大 XLSX 流式读取。

本文档面向 **5.0.1** 版本，适用于托管在 **GitHub Pages**（例如站点根目录指向 `docs/`）时的版本化说明页。

- 设计与实现审查（工程向）：[5.0.1 审查报告](./risks.md)

---

## 环境要求

- **JDK 17**（与当前工程 toolchain 一致）。
- 运行时需具备 **Vert.x Core** 提供的 `io.vertx.core.Future` 等 API（本库通过 `api` 依赖传递 **keel-core**，一般只需声明
  `keel-poi` 即可）。
- 许可证：**GPL-3.0**（见仓库与已发布 POM）。

---

## 在项目中引入依赖

**Maven：**

```xml

<dependency>
    <groupId>io.github.sinri</groupId>
    <artifactId>keel-poi</artifactId>
    <version>5.0.1</version>
</dependency>
```

**Gradle（Kotlin DSL）：**

```kotlin
implementation("io.github.sinri:keel-poi:5.0.1")
```

> 若你本地仍为 `5.0.1-SNAPSHOT`，请将版本号换成实际发布坐标；正式发布后以 Maven Central 或项目 README 为准。

---

## 核心概念

| 类型                                             | 作用                                               |
|------------------------------------------------|--------------------------------------------------|
| `KeelSheets`                                   | 工作簿入口：打开已有文件或创建新工作簿，用完后在 `useSheets` 的回调链末尾自动关闭。 |
| `SheetsOpenOptions` / `SheetsCreateOptions`    | 打开与新建时的选项（文件或流、是否 XLSX、公式求值器、大文件流式、XLSX 流式写入等）。  |
| `KeelSheet`                                    | 单张工作表的读写：矩阵/模板化矩阵、按行遍历、写入单元格、读取图片等。              |
| `KeelSheetMatrix` / `KeelSheetTemplatedMatrix` | 表头 + 数据行的内存模型；模板化矩阵可按列名访问。                       |
| `KeelCsvReader` / `KeelCsvWriter`              | CSV 按行解析与写入（支持分隔符、字符集、`RFC 4180` 风格转义）。          |

公共 API 多使用 **`Future`** 表达异步边界；**阻塞式**方法（如 `readAllRowsToMatrix`）适合在 Worker 线程或简单脚本中使用。

---

## Excel：打开已有工作簿

使用 **`KeelSheets.useSheets(SheetsOpenOptions, usage)`**：在 `usage` 中返回的 `Future` 完成后，库会关闭工作簿。

**从文件打开：**

```java
import io.github.sinri.keel.integration.poi.excel.*;
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

**从输入流打开：**

- 若 **`setUseXlsx(null)`**（默认）：库会先把流 **读入字节数组** 再尝试 XLSX / XLS 判别，**不适合极大文件**。
- **`setInputStream` 与 `setFile` 互斥**：设置其中一个会清空另一个（5.0.1 行为），避免隐式优先级困扰。

**超大 XLSX 流式读取**（基于 [excel-streaming-reader](https://github.com/pjfanning/excel-streaming-reader)）：

1. 调用 **`SheetsOpenOptions.declareReadingVeryLargeExcelFiles()`
   **（建议在进程启动时执行一次），按 Javadoc 调整 POI 临时文件等行为。
2. 使用 **`setHugeXlsxStreamingReaderBuilder(handler)`** 配置 `StreamingReader.Builder`，再通过文件或流打开。

流式模式下可逐行访问当前行单元格，**不能像普通 XSSF 一样随机跳行**。

---

## Excel：创建新工作簿并保存

**`KeelSheets.useSheets(SheetsCreateOptions, usage)`** 在结束后同样会关闭工作簿。

```java
SheetsCreateOptions opts = new SheetsCreateOptions()
        .setUseXlsx(true)
        .setUseStreamWriting(true)  // 仅对 XLSX 生效：内部使用 SXSSF，适合大批量写行
        .setWithFormulaEvaluator(false);

KeelSheets.

useSheets(opts, sheets ->{
KeelSheet w = sheets.generateWriterForSheet("Data");
KeelSheetMatrix m = new KeelSheetMatrix();
    m.

setHeaderRow(List.of("A", "B"));
        m.

addRow(List.of("1", "2"));
        w.

writeMatrix(m);
    sheets.

save("/path/to/out.xlsx");
    return Future.

succeededFuture();
});
```

- **`setUseStreamWriting(true)`** 在 **`setUseXlsx(false)`**（HSSF）时会被 **静默忽略**，仅 XLSX 下有意义。

---

## Excel：读取为矩阵或模板化矩阵

**默认约定**（`readAllRowsToMatrix()` / `readAllRowsToTemplatedMatrix()`）：

1. **第一行作为表头**（表头前的行会被丢弃，见带 `headerRowIndex` 的重载）。
2. **列数**：若 `maxColumns <= 0`，在表头行上用 **首个非空列范围** 自动推断列数。
3. **行过滤**：默认排除「全空单元格」行（见下文 **`SheetRowFilter`**）。

重载签名示例：

```java
// headerRowIndex：表头所在行号（0-based）
// maxColumns：>0 为固定列数；<=0 为自动检测
// sheetRowFilter：可 null 表示不过滤
KeelSheetMatrix matrix = sheet.readAllRowsToMatrix(0, 0, SheetRowFilter.toExcludeEmptyRows());
```

**异步批量读取**（需 **`Keel`** 实例，按批处理 POI `Row`）：

- `readAllRowsToMatrixAsync(keel, ...)`
- `readAllRowsToTemplatedMatrixAsync(keel, ...)`

适用于已在 Vert.x/Keel 工作线程上下文中、希望避免长时间阻塞调度线程的场景。

**矩阵 → 模板化矩阵**（按表头列名绑定）：

```java
KeelSheetTemplatedMatrix templated = matrix.transformToTemplatedMatrix();
```

表头为空时会抛异常（列未定义）。

**自定义行模型（5.0.1）**

`KeelSheetMatrix.getRowIterator(Function<List<String>, R> rowFactory)` 使用工厂函数构造你的 `KeelSheetMatrixRow` 子类；*
*已不再提供**基于 `Class` + 反射的迭代器重载。

**`KeelSheetMatrixRow` 数值 API**

`readValueToInteger` / `readValueToLong` 等在 **无法解析或越界** 时返回 **`null`**（含 `NumberFormatException` 场景）。

---

## Excel：写入矩阵与原始行

- **`writeMatrix(KeelSheetMatrix)`**：有表头则第 0 行写表头，数据从第 1 行起；无表头则只写数据行。
- **`writeTemplatedMatrix`** / 对应的 `*Async`：写入模板化矩阵。
- **`writeAllRows(List<List<String>> rowData, sinceRowIndex, sinceCellIndex)`**：低开销直接按行列下标写字符串。

---

## Excel：公式与图片

- **公式**：在 **`SheetsOpenOptions` / `SheetsCreateOptions`** 上 **`setWithFormulaEvaluator(true)`**，并在 *
  *`generateReaderForSheet`** 时使用默认或 `parseFormulaCellToValue == true`，可将公式单元格按求值结果读出（具体行为见 POI
  `FormulaEvaluator`）。
- **图片**：**`KeelSheet.getPictures()`** 返回 **`KeelPictureInSheet`
  ** 列表，用于从绘图层提取内嵌图片（实现同时覆盖 XLS / XLSX 路径）。

---

## 行过滤：`SheetRowFilter`（5.0.1）

内置 **排除全空行**：

```java
SheetRowFilter filter = SheetRowFilter.toExcludeEmptyRows();
```

- **推荐**使用 **`shouldExcludeRow(rawRow)`**（返回 `true` 表示 **丢弃** 该行）。
- 旧名 **`shouldThrowThisRawRow`** / **`toThrowEmptyRows`** 在 **5.0.1** 起标记为 **`@Deprecated`**，语义不变，建议迁移到新命名。

可自实现接口，按业务决定是否丢弃行。

---

## CSV：读取

推荐使用静态方法 **`KeelCsvReader.read(InputStream, Charset, separator, readFunc)`**，由库在结束时关闭读取器。

```java
KeelCsvReader.read(
        inputStream,
        StandardCharsets.UTF_8,
    ",",
        reader ->{
        try{
CsvRow row;
            while((row =reader.

next())!=null){
        // 5.0.1：可按行遍历单元格
        row.

stream().

forEach(cell ->{ /* ... */ });
        // 或 List<CsvCell> cells = row.getCells();
        }
        }catch(
IOException e){
        return Future.

failedFuture(e);
        }
                return Future.

succeededFuture();
    }
            );
```

解析支持带引号字段、分隔符与换行等常见 CSV 变体（实现以库内 `consumeOneLine` 逻辑为准）。

---

## CSV：写入

- **`KeelCsvWriter.write(OutputStream, separator, Charset, writeFunc)`** 或仅 **`write(OutputStream, writeFunc)`
  **（默认 UTF-8、逗号分隔）。
- 逻辑行由 **`writeCell` 多次调用 + `writeRowEnding()`** 组成；或使用 **`blockWriteRow(List)`** 一次性写一行。
- 行分隔符为 **`CRLF`（`\r\n`）**，以贴近 RFC 4180。

---

## 常见问题与注意点

1. **`useSheets` 回调必须返回 `Future`**，以便链式 `eventually` 关闭工作簿；不要在回调外长期持有已关闭的 `Workbook`。
2. **流式 XLSX** 与 **普通 XSSF
   ** 的能力差异（随机访问行、共享字符串等）需按 [excel-streaming-reader](https://github.com/pjfanning/excel-streaming-reader) 文档评估。
3. **SXSSF 写入** 仅缓存滑动窗口内的行，极大量写入时需了解 POI 行窗口与磁盘刷写策略。
4. 更细的设计权衡与已知测试缺口见 **[审查报告](./risks.md)**。

---

## 相关链接

- 源码与议题：<https://github.com/sinri/keel-poi>
- 上游 Keel Core：<https://github.com/sinri/keel-core>
