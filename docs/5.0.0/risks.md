# keel-poi 项目审查报告

各风险条目下方的 **处理状态** 表示截至文档更新时在代码库中的处置结论（对应提交大致为：高危 `85e90cb`、严重 `5b70c93`、中等 `4109dd4`、低危 `ebead0a`，具体以 `git log` 为准）。

---

## 一、严重 Bug（Critical）

### 1. KeelSheets.save — FileOutputStream 资源泄漏

文件: `KeelSheets.java`（`save(File)` 原实现）

`new FileOutputStream(file)` 创建后直接传入 `save(OutputStream)`，流未关闭；若 `autoWorkbook.write()` 抛异常，流也不会关闭。应使用 try-with-resources。

**处理状态：** 已修复。`save(File)` 使用 `try (FileOutputStream fos = …)` 并捕获 `IOException`。

---

### 2. KeelSheets.useSheets — andThen 中的 close 结果被丢弃

文件: `KeelSheets.java`（`useSheets` 两处）

`close(promise)` 写入的 `Promise` 未与返回的 `Future` 组合，关闭失败易被吞掉；`andThen` 亦不等待关闭完成。

**处理状态：** 已修复。两处均改为 `Future.eventually`，在业务 `Future` 完成后执行 `close(promise)` 并等待其 `Future`。

---

### 3. KeelSheets 自动检测格式 — InputStream 被消耗后无法回退

文件: `KeelSheets.java`（`useSheets` 打开选项分支）

先 `XSSFWorkbook(inputStream)` 失败再同流构造 `HSSFWorkbook` 时，流可能已前移至不可用状态。

**处理状态：** 已修复。在 `isUseXlsx() == null` 时先 `readAllBytes()`，再对 `ByteArrayInputStream` 分别尝试 XSSF / HSSF。

---

### 4. KeelSheet.writeMatrixAsync — 数据写入逻辑错误

Lambda 收到当前批次 `rawRows`，但原实现每批仍写入 `matrix.getRawRowList()` 全量，导致重复写入与越界。

**处理状态：** 已修复。改为 `writeAllRows(rawRows, rowIndexRef.get(), 0)`，并与 `rowIndexRef.addAndGet(rawRows.size())` 一致。

---

### 5. KeelSheet.writeTemplatedMatrix — 行索引未递增

`forEach` 内始终使用同一 `rowIndexRef.get()`，未 `incrementAndGet`，多行叠在同一行。

**处理状态：** 已修复。每写一行后 `rowIndexRef.incrementAndGet()`。

---

### 6. KeelSheet.dumpCellToString — BOOLEAN/ERROR 等误用 getStringCellValue()

非 NUMERIC / FORMULA 分支曾统一 `getStringCellValue()`，对 BOOLEAN、ERROR 等会抛 `IllegalStateException`。`autoDetectNonBlankColumnCountInOneRow` 亦曾误用字符串读取。

**处理状态：** 已修复。`dumpCellToString` 按 `CellType` 分支处理；无 evaluator 的 FORMULA 走 `getCachedFormulaResultType()`。列数检测改为基于 `Row.getLastCellNum()`（与下述 #18 一并解决截断问题）。

---

## 二、高危问题（High）

### 7. KeelSheets FormulaEvaluator 绑定到错误的 Workbook

创建模式下用 XSSF 建 `FormulaEvaluator` 后又将 `autoWorkbook` 换为 `SXSSFWorkbook`，求值器仍指向旧簿。

**处理状态：** 已修复。`formulaEvaluator` 在替换为 SXSSF 后按当前 `autoWorkbook` 重建。

---

### 8. KeelCsvReader.consumeOneLine — 递归栈溢出

跨多物理行的引用字段原递归 `consumeOneLine`，深层递归可 `StackOverflowError`。

**处理状态：** 已修复。改为 `while` 循环读下一行，直至引号闭合或 EOF。

---

### 9. KeelCsvReader — 多字符分隔符不工作

逐字与 `separator` 比较，多字符分隔符无法匹配。

**处理状态：** 已修复。使用 `regionMatches` 匹配整段分隔符，并推进索引；空 `separator` 不参与匹配。

---

### 10. KeelSheet.readRow 返回 null 但未标注

`sheet.getRow(i)` 可返回 null，`readRawRow` 传入 `dumpRowToRawRow` 曾导致 NPE。

**处理状态：** 已修复。`readRow` 标 `@Nullable Row`；`dumpRowToRawRow` 对 `row == null` 按空白行（给定列数）处理并仍走行过滤器。

---

### 11. KeelPictureInSheet HSSF 尺寸计算逻辑错误

曾用 `getDx1()/getDy1()` 当宽高，数值严重偏小。

**处理状态：** 已修复。与 XSSF 一致使用 `HSSFPicture.getImageDimension()`。

---

## 三、中等问题（Medium）

### 12. KeelCsvWriter.writeCell — 多余的尾部分隔符

每格后都写 `separator`，与 `writeRowEnding` 组合会产生尾逗号与幽灵列。

**处理状态：** 已修复。首格前不加分隔符，后续格前写分隔符，末格后无尾随分隔符。

---

### 13. KeelCsvWriter.quote — replaceAll 与语义

引号加倍曾用 `replaceAll`。

**处理状态：** 已修复。改为 `String.replace`；字段中含 `\r` 时也按需加引号。

---

### 14. KeelCsvWriter — 行结束符

原使用 `\n`，RFC 4180 推荐 CRLF。

**处理状态：** 已修复。统一使用 `\r\n`（`blockWriteRow`、`writeRowEnding`、补全未结束行的换行）。

---

### 15. KeelCsvReader / KeelCsvWriter 静态方法 — eventually 与异常

业务失败时若 `close` 再失败，`eventually` 可能只暴露关闭异常。

**处理状态：** 已修复。改为 `compose(成功时关闭, 失败时关闭)`；关闭失败时对业务异常 `addSuppressed(closeErr)` 后仍 `failedFuture(业务异常)`。

---

### 16. KeelSheetMatrix 迭代器 — AtomicInteger 误导

原子类暗示线程安全，但迭代与底层 `ArrayList` 均非并发安全。

**处理状态：** 已修复。改为 `int` 游标，并在 `RowReaderIterator` JavaDoc 标明非线程安全。

---

### 17. KeelSheetMatrixRowTemplateImpl — 重复表头静默覆盖

重复列名在 `LinkedHashMap` 中覆盖，前列名丢失。

**处理状态：** 已修复。构造时检测重复列名并 `IllegalArgumentException`；接口 `create` 文档已说明。

---

### 18. autoDetectNonBlankColumnCountInOneRow — 遇到首个空列即停止

中间空列会导致列数被截断。

**处理状态：** 已修复（与 #6 同批）。实现改为返回 `row.getLastCellNum()`（无单元格时为 0），不再在首个空/null 格处提前 `break`。（注：若报告撰写时行号指向旧实现逻辑，以当前 `KeelSheet` 源码为准。）

---

## 四、低危 / 代码风格（Low）

### 19. KeelSheetDrawing 构造函数 — 冗余 null 赋值

曾对 `ValueBox` 重复 `setValue(null)` 等死代码。

**处理状态：** 已修复。仅在存在 `DrawingPatriarch` 时写入；读取侧用 `isValueAlreadySet()` 区分「从未设置」与「设为 null」。

---

### 20. KeelPictureInSheet.getData() — 暴露内部 byte[]

**处理状态：** 已修复。返回 `data.clone()`，JavaDoc 说明防御性拷贝。

---

### 21. 注释中残留 commented-out 代码

`KeelSheetMatrixRow`、`KeelSheetTemplatedMatrixImpl` 等。

**处理状态：** 已删除相关注释代码块。

---

### 22. 类名过长且相似 — KeelSheetMatrixTemplatedRow vs KeelSheetMatrixRowTemplate

**处理状态：** 已缓解（未重命名类型）。在两者的类/接口 JavaDoc 中明确：其一描述整表列结构，其一描述单行按模板的访问视图，降低误用成本。

---

### 23. build.gradle.kts — sonatype 凭据在配置阶段强制解析

`by project` 缺失属性可导致无关任务（如 `compileJava`）在配置期失败。

**处理状态：** 已修复。移除顶层 `sonatypeUsername`/`Password` 的 `by project`；JReleaser 块内使用 `findProperty(…).orEmpty()`。

---

### 24. Excel 模块缺少测试

**处理状态：** 已修复（基础覆盖）。新增 `KeelExcelCoreTest`：`dumpCellToString`、矩阵写入与 `save(File)` 及回读校验；CSV 侧另有既有测试与其它批次增补用例。模板矩阵等仍可按需加宽覆盖。

---

## 总结

| 严重程度 | 数量 | 关键主题 | 处理概况 |
|----------|------|----------|----------|
| 严重 (Critical) | 6 | 流与异步关闭、格式探测、矩阵异步/模板写、单元格类型 | 均已修复 |
| 高危 (High) | 5 | FormulaEvaluator、CSV 解析、可空行、图片尺寸 | 均已修复 |
| 中等 (Medium) | 7 | CSV 分隔符与 CRLF、异常链、迭代器、重复表头、列宽检测 | 均已修复 |
| 低危 (Low) | 6 | Drawing 清理、byte[] 拷贝、注释、文档区分、Gradle、测试 | #22 为文档缓解，其余已修复 |

历史优先级说明（原报告结论）：#4、#5 曾是最易直接导致数据损坏的项；#3、#1 为资源与流正确性——**当前代码库中对应项已按上表处理**。
