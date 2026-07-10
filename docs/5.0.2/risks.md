# keel-poi 5.0.2 代码审查报告

基于 `dev-5.0.2` 分支当前代码、既有 `5.0.0` / `5.0.1` 审查记录，以及 `./gradlew test` 的验证结果进行复审。本文保留初始发现，并记录当前分支的处理状态。

验证结果：
`./gradlew test` 通过。5.0.2 已补充同步异常自动关闭、行过滤迭代、模板化缺失列、矩阵防御性拷贝、XLSX 流式读取和图片提取等边界路径的回归测试。

---

## 一、Bug / 潜在运行时错误

### #1 [高] 自动关闭包装器在用户回调同步抛异常时不会关闭资源

**文件：**

- `src/main/java/io/github/sinri/keel/integration/poi/csv/KeelCsvReader.java:70-90`
- `src/main/java/io/github/sinri/keel/integration/poi/csv/KeelCsvWriter.java:71-91`
- `src/main/java/io/github/sinri/keel/integration/poi/excel/KeelSheets.java:89-170`
- `src/main/java/io/github/sinri/keel/integration/poi/excel/KeelSheets.java:183-210`

这些方法声明会自动管理 Reader / Writer / Workbook 生命周期，但关闭逻辑只挂在 `readFunc.apply(...)`、
`writeCsvFunc.apply(...)`、`usage.apply(...)` 返回的 `Future` 上。

如果用户回调在返回 `Future` 前同步抛出异常，已经创建的资源不会进入后续 `compose` / `eventually` 分支，导致流或工作簿未关闭。对默认
`SheetsCreateOptions` 而言，底层可能是 `SXSSFWorkbook`；POI 5.4.1 的 `close()` 会清理临时文件，因此漏关会直接破坏流式写入的临时文件清理语义。

**建议：** 在调用用户回调处用
`try/catch` 包住同步异常；异常路径中执行关闭，并保留原始异常，关闭失败时作为 suppressed 或按现有策略返回。也可抽出小型
`closeAfter(Future<T>)` / `closeOnSyncFailure(...)` 辅助方法，避免三处重复。

**处理状态：** 已修复。`KeelCsvReader.read(...)`、`KeelCsvWriter.write(...)` 和 `KeelSheets.useSheets(...)` 均把用户回调包入
`Future.succeededFuture().compose(...)`，同步异常会转为失败
`Future` 并进入关闭分支。已补充 CSV Reader、CSV Writer 与 Excel 工作簿自动关闭回归测试。

---

### #2 [中] 行过滤器会让 `getRawRowIterator` 返回 null，后续矩阵迭代器可能产生隐藏 NPE

**文件：** `src/main/java/io/github/sinri/keel/integration/poi/excel/KeelSheet.java:274-287`, `682-693`, `707-722`

`dumpRowToRawRow(...)` 在行被 `SheetRowFilter` 丢弃时返回 `null`，`getRawRowIterator(...)` 的 `next()` 也直接返回这个
`null`。但方法签名是 `Iterator<List<String>>`，调用者自然会认为 `next()` 返回非空行。

更明显的是 `getMatrixRowIterator(...)` 与 `getTemplatedMatrixRowIterator(...)` 直接把
`rawRowIterator.next()` 传给行对象创建逻辑。只要过滤器丢弃某行，就可能创建包含 `null` rawRow 的
`KeelSheetMatrixRow`，或创建行为不明确的模板化行，错误会延迟到后续读取列值时才暴露。

**建议：** 迭代器层面跳过被过滤行，使 `next()` 永远返回实际行；或将方法签名改为可空语义并同步调整矩阵迭代器。优先建议跳过过滤行，因为这更符合
`SheetRowFilter` 的“排除行”语义。

**处理状态：** 已修复。
`getRawRowIterator(...)` 在迭代器内部预取下一条未被过滤的实际行；若过滤器排除当前行，会继续读取下一行。矩阵行迭代器和模板化行迭代器复用该行为。已补充三类迭代器跳过空行的回归测试。

---

### #3 [中] `KeelSheetMatrixTemplatedRowImpl.getColumnValue(String)` 与按索引访问行为不一致

**文件：** `src/main/java/io/github/sinri/keel/integration/poi/excel/entity/KeelSheetMatrixTemplatedRowImpl.java:53-72`

`getColumnValue(int)` 已在索引超过 `rawRow.size()` 时返回空字符串，适配 Excel 尾部空列常被截断的情况。但
`getColumnValue(String)` 找到列索引后直接 `rawRow.get(columnIndex)`，没有复用按索引访问逻辑。

当模板包含列 `C`，实际数据行只有 `A/B` 两列时：

```java
row.getColumnValue(2);   // 返回 ""
row.getColumnValue("C"); // 抛 IndexOutOfBoundsException
```

这会让模板化行的两套访问方式出现不一致，也会重新引入 5.0.1 审查中“尾部空列截断导致越界”的同类问题。

**建议：** `getColumnValue(String)` 在解析列索引后调用 `getColumnValue(columnIndex)`。对于不存在的列名，可继续抛异常，但建议改成带列名上下文的
`IllegalArgumentException`。

**处理状态：** 已修复。`getColumnValue(String)` 解析列索引后复用 `getColumnValue(int)`，列存在但数据行尾部缺失时返回空字符串；列名不存在时抛出包含列名上下文的
`IllegalArgumentException`。已补充回归测试。

---

## 二、封装与可变性问题

### #4 [中] 矩阵返回不可变外层列表，但内部行列表仍可被外部修改

**文件：**

- `src/main/java/io/github/sinri/keel/integration/poi/excel/entity/KeelSheetMatrix.java:42-99`
- `src/main/java/io/github/sinri/keel/integration/poi/excel/entity/KeelSheetTemplatedMatrixImpl.java:38-83`
- `src/main/java/io/github/sinri/keel/integration/poi/excel/entity/KeelSheetMatrixTemplatedRowImpl.java:81-83`

5.0.1 已把若干 getter 改为 `Collections.unmodifiableList(...)`，但目前只保护了外层列表。调用者仍可通过以下方式修改矩阵内部状态：

```java
matrix.getRawRow(0).set(0, "changed");
matrix.getRawRowList().get(0).set(0, "changed");
templatedMatrix.getRawRows().get(0).set(0, "changed");
```

此外，`addRow(...)` / `addRawRow(...)` 直接保存调用者传入的 `List` 引用，调用者在添加后继续修改原列表，也会改变矩阵内部状态。

**建议：** 写入内部状态时使用 `List.copyOf(...)` 或
`new ArrayList<>(...)` 做防御性拷贝；返回嵌套列表时同时保护每一行。若需要保留可变行能力，应在 JavaDoc 中明确这是有意设计。

**处理状态：** 已修复。`KeelSheetMatrix`、`KeelSheetTemplatedMatrixImpl` 和
`KeelSheetMatrixTemplatedRowImpl` 在写入内部状态时复制行列表，返回嵌套列表时同时保护外层与内层列表。已补充防御性拷贝与只读视图回归测试。

---

### #5 [低] `KeelSheetMatrixRowTemplateImpl` 直接持有外部表头列表，后续修改会破坏模板一致性

**文件：** `src/main/java/io/github/sinri/keel/integration/poi/excel/entity/KeelSheetMatrixRowTemplateImpl.java:29-40`,
`74-76`

模板构造时用传入的 `headerRow` 构建 `headerMap`，但字段仍直接保存原列表引用。若调用方在创建模板后修改原列表，
`getColumnNames()` / `getColumnName(i)` 看到的是新列表，而 `getColumnIndex(name)` 仍来自旧 `headerMap`，两者会不一致。

**建议：** 构造函数中复制表头列表，并让 `getColumnNames()` 返回不可变视图或副本。

**处理状态：** 已修复。`KeelSheetMatrixRowTemplateImpl` 构造时复制并归一化表头，
`getColumnNames()` 返回只读视图。已补充表头防御性拷贝、只读列名和空表头项归一化测试。

---

## 三、测试覆盖

### #6 [中] 5.0.1 已记录的测试缺口仍未完全补齐

**涉及范围：**

- 模板化矩阵：`KeelSheetTemplatedMatrix`、`KeelSheetMatrixTemplatedRow`
- 行过滤器：`SheetRowFilter`
- 图片提取：`KeelSheetDrawing`、`KeelPictureInSheet`
- XLSX 流式读取：`SheetsOpenOptions.setHugeXlsxStreamingReaderBuilder(...)`
- 同步异常下的自动关闭路径

初始复审时 `./gradlew test` 通过，但这些路径缺少回归测试。#1、#2、#3 都属于原有测试不容易捕获的边界路径。

**建议：** 最少补充以下测试：

- 自动关闭包装器：用户回调同步抛异常时仍关闭资源；
- `SheetRowFilter` 丢弃行时，公开迭代器不返回 `null`；
- 模板化行按索引和按列名读取尾部缺失列时行为一致；
- 模板/矩阵对外部可变列表的防御性拷贝。

**处理状态：** 已处理核心缺口。当前已补充自动关闭包装器同步异常、
`SheetRowFilter` 迭代跳过、模板化行尾部缺失列、防御性拷贝、XLSX 流式读取和图片提取测试。仍可在后续版本继续扩展 XLS/HSSF 图片、更多公式类型和异常关闭组合测试。

---

## 四、发布前检查

### #7 [低] 5.0.2 版本升级变更已在工作区，但尚未形成发布记录

**文件：** `gradle.properties`

当前未提交修改把版本改为 `5.0.2`，并升级了 `keelCoreVersion`、`keelTestVersion`、`excelStreamingReaderVersion`。这与
`dev-5.0.2` 目标一致。

**建议：** 在处理本报告风险后，为 5.0.2 增加发布说明，至少记录依赖升级、兼容性影响、修复项和测试结果。

**处理状态：** 已处理。已新增 `docs/5.0.2/index.md`，并更新 `docs/index.md` 的 Latest Version 与版本入口。

---

## 总结

| 严重程度 | 数量 | 主题                           |
|------|----|------------------------------|
| 高    | 1  | 自动关闭包装器同步异常漏关资源              |
| 中    | 4  | 行过滤器迭代器、模板化行读取一致性、矩阵可变性、测试缺口 |
| 低    | 2  | 模板表头防御性拷贝、5.0.2 发布记录         |

当前分支已处理 #1 至 #7。5.0.2 的主要用户可见变化是资源自动关闭更可靠、行过滤迭代器不再返回
`null`、模板化列访问行为更一致，以及矩阵/模板对象对外部可变列表的防御更严格。
