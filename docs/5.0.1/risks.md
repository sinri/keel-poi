# keel-poi 5.0.1 设计与实现审查报告

基于对 5.0.0 审查报告（24项，均已处理）之后的全量代码复审。本报告聚焦 5.0.0 修复后残留的问题和新发现的设计缺陷。

---

## 一、Bug / 潜在运行时错误

### #1 [中] KeelSheetMatrixRow.readValueToInteger / readValueToLong — 未捕获 NumberFormatException

**文件:** `KeelSheetMatrixRow.java:72-95`

`readValueToInteger` 和 `readValueToLong` 只捕获 `ArithmeticException`（来自 `intValueExact`/`longValueExact`），但 `readValueToBigDecimal` 内的 `new BigDecimal(string)` 会在非数值字符串时抛出 `NumberFormatException`，该异常未被捕获。

JavaDoc 声称"如果转换失败则返回 null"，但实际上非数字字符串会直接抛异常。

**建议:** catch 块改为 `NumberFormatException | ArithmeticException`，或统一为 `catch (Exception e)`。

**处理状态：** 已修复。catch 块改为 `NumberFormatException | ArithmeticException`。

---

### #2 [中] KeelSheetMatrixTemplatedRowImpl.toJsonObject — rawRow 长度不足时 IndexOutOfBoundsException

**文件:** `KeelSheetMatrixTemplatedRowImpl.java:90-97`

```java
for (int i = 0; i < columnNames.size(); i++) {
    x.put(columnNames.get(i), getColumnValue(i)); // → rawRow.get(i)
}
```

若实际数据行的列数少于模板定义的列数（Excel 中常见：尾部空列被 POI 截断），`rawRow.get(i)` 将抛出 `IndexOutOfBoundsException`。

**建议:** 越界时返回空字符串 `""`，或在 `addRawRow` 时填充至模板列数。

**处理状态：** 已修复。`getColumnValue(int)` 在索引越界时返回空字符串。

---

### #3 [中] KeelPictureInSheet(XSSFPicture) — getImageDimension() 在 anchor 为 null 时仍被调用

**文件:** `KeelPictureInSheet.java:34-54`

XSSF 构造函数中，`getImageDimension()` 在 anchor null-check 分支之外（第44行）。当 `getClientAnchor() == null` 时，`getImageDimension()` 仍然会被执行。对比 HSSF 分支，尺寸获取在 anchor 非空分支内部。

**建议:** 将第44-46行移入 anchor != null 的 if 分支内，else 分支设 `width = -1; height = -1`，与 HSSF 保持一致。

**处理状态：** 已修复。`getImageDimension()` 移入 anchor 非空分支内，else 设 `width = -1; height = -1`。

---

### #4 [低] KeelSheetMatrix.RowReaderIterator.next() — 未检查 hasNext，异常语义不清

**文件:** `KeelSheetMatrix.java:206-209`

`next()` 在迭代器耗尽时会抛出 `IndexOutOfBoundsException`（来自 `rows.get(nextIndex)`），而非标准的 `NoSuchElementException`。

**建议:** 在 `next()` 开头添加 `if (!hasNext()) throw new NoSuchElementException();`。

**处理状态：** 已修复。添加 `hasNext()` 检查，抛出 `NoSuchElementException`。

---

### #5 [低] KeelSheetMatrixRow — 无越界上下文

**文件:** `KeelSheetMatrixRow.java:37-39`

`readValue(int i)` 直接调用 `rawRow.get(i)`，越界时只有裸的 `IndexOutOfBoundsException`，无法知道请求了第几列、行有几列。

**建议:** 可用 `Objects.checkIndex` 或自定义消息包装。影响低，视需求而定。

**处理状态：** 已修复。添加越界检查，提供列索引和行列数的上下文信息。

---

## 二、设计问题

### #6 [中] 内部可变列表直接暴露 — 外部可篡改矩阵状态

**涉及文件:**
- `KeelSheetMatrix.getHeaderRow()` 返回内部 `ArrayList`（第63行）
- `KeelSheetMatrix.getRawRowList()` 返回内部 `rows`（第98-99行）
- `KeelSheetTemplatedMatrixImpl.getRawRows()` 返回内部 `rawRows`（第39-41行）
- `KeelSheetMatrixTemplatedRowImpl.getRawRow()` 返回内部 `rawRow`（第78-79行）

调用者可以任意修改这些列表，破坏矩阵的内部状态。例如 `matrix.getHeaderRow().clear()` 会清空表头。

**建议:** 返回 `Collections.unmodifiableList(...)` 或 `List.copyOf(...)`。

**处理状态：** 已修复。上述四处均改为返回 `Collections.unmodifiableList(...)`。

---

### #7 [中] SheetsOpenOptions 允许同时设置 File 和 InputStream — 无冲突校验

**文件:** `SheetsOpenOptions.java`

用户可以同时调用 `setFile(...)` 和 `setInputStream(...)`。`KeelSheets.useSheets` 按 InputStream 优先处理，但这种隐式优先级未在文档中说明，也没有在设置时校验互斥。

**建议:** 在 `setFile` 时将 `inputStream` 置 null（反之亦然），或在 `useSheets` 入口校验两者互斥。

**处理状态：** 已修复。`setFile` 时清空 `inputStream`，`setInputStream` 时清空 `file`。

---

### #8 [低] SheetsCreateOptions.useStreamWriting 对 HSSFWorkbook 静默无效

**文件:** `SheetsCreateOptions.java:14`, `KeelSheets.java:189-199`

`useStreamWriting` 默认 `true`，但 `SXSSFWorkbook` 仅适用于 XSSF。当 `useXlsx = false` 时，`useStreamWriting = true` 被静默忽略。

**建议:** 在 Javadoc 中说明仅对 XLSX 生效，或在 `setUseStreamWriting(true)` 时校验 `useXlsx`。

**处理状态：** 已修复。在 `isUseStreamWriting`/`setUseStreamWriting` 的 Javadoc 中说明仅对 XLSX 有效。

---

### #9 [低] SheetsOpenOptions.isWithFormulaEvaluator 为 package-private 但 SheetsCreateOptions 中为 public

**文件:** `SheetsOpenOptions.java:39` vs `SheetsCreateOptions.java:61`

同名方法可见性不一致，对使用者造成困惑。

**建议:** 统一为 `public`，或提供明确的 `isWithFormulaEvaluator()` getter。

**处理状态：** 已修复。`SheetsOpenOptions.isWithFormulaEvaluator()` 改为 `public` 并补充 Javadoc。

---

### #10 [低] SheetRowFilter 方法名 shouldThrowThisRawRow 语义不清

**文件:** `SheetRowFilter.java:40`

"throw" 容易被理解为 "抛出异常"。实际含义是 "丢弃此行"。

**建议:** 重命名为 `shouldExcludeRow` 或 `shouldSkipRow`（需考虑向后兼容）。

**处理状态：** 已修复。添加 `shouldExcludeRow` 默认方法和 `toExcludeEmptyRows` 工厂方法，旧方法标记 `@Deprecated(since = "5.0.1")`。

---

### #11 [低] KeelSheetDrawing 构造函数 — 无类型检查的强制转型

**文件:** `KeelSheetDrawing.java:40,45`

```java
drawingForXlsxValueBox.setValue((XSSFDrawing) x);
drawingForXlsValueBox.setValue((HSSFPatriarch) x);
```

虽然 `sheetsReaderType` 分支逻辑上保证了类型匹配，但未使用 `instanceof` 做防御性检查。如果 POI 版本升级改变了返回类型，将产生无上下文的 `ClassCastException`。

**建议:** 加 `instanceof` 判断，失败时抛带有描述信息的异常。

**处理状态：** 已修复。使用 `instanceof` 模式匹配替代强制转型。

---

## 三、KeelPictureInSheet 残留

### #12 [低] 注释代码残留

**文件:** `KeelPictureInSheet.java:80`

```java
// int format = pictureData.getFormat();// HSSF specific
```

5.0.0 审查 #21 声称已删除残留注释代码块，但此处仍保留。

**处理状态：** 已修复。已删除该注释行。

---

## 四、测试覆盖

### #13 [中] 模板化矩阵子系统零测试

以下类/接口没有任何测试覆盖：
- `KeelSheetTemplatedMatrix` / `KeelSheetTemplatedMatrixImpl`
- `KeelSheetMatrixTemplatedRow` / `KeelSheetMatrixTemplatedRowImpl`
- `KeelSheetMatrix.transformToTemplatedMatrix()`
- `KeelSheetMatrixTemplatedRowImpl.toJsonObject()`

这些是用户常用的高层 API，缺少测试意味着 #2 等 bug 无法被自动检测到。

**处理状态：** 待补充测试。

---

### #14 [低] KeelSheetMatrixRow 数值转换无测试

`readValueToInteger`、`readValueToLong`、`readValueToDouble`、`readValueToBigDecimalStrippedTrailingZeros` 均无测试。 #1 所述的 `NumberFormatException` 遗漏即因此未被发现。

**处理状态：** 待补充测试。

---

### #15 [低] SheetRowFilter、KeelSheetDrawing、KeelPictureInSheet 无测试

行过滤器和图片提取功能无测试覆盖。

**处理状态：** 待补充测试。

---

### #16 [低] 流式读取路径 (XLSX_STREAMING) 无测试

`SheetsOpenOptions.setHugeXlsxStreamingReaderBuilder` 所触发的流式读取路径从未在测试中执行。

**处理状态：** 待补充测试。

---

## 五、代码风格 / 杂项

### #17 [低] KeelSheetMatrix.getRowIterator() 使用反射实例化

**文件:** `KeelSheetMatrix.java:162-171`

`RowReaderIterator(Class<R>)` 构造器通过 `rClass.getConstructor(List.class).newInstance(strings)` 反射创建行对象。这在 JPMS 下对未 exports 的包会失败，且反射带来不必要的性能开销和异常包装 (`InvocationTargetException`)。

重载版 `getRowIterator()` （第137-140行）已正确使用 lambda，建议统一移除反射版本，改为要求调用者提供 `Function<List<String>, R>` 工厂。

**处理状态：** 已修复。移除 `RowReaderIterator(Class<R>)` 构造器和 `getRowIterator(Class<R>)` 方法，改为 `getRowIterator(Function<List<String>, R>)` 工厂函数版本。

---

### #18 [低] CsvRow 缺少迭代和批量访问 API

**文件:** `CsvRow.java`

仅有 `getCell(int)` 和 `size()`，无 `getCells()`、`iterator()`、`stream()` 或 `toList()` 方法。用户需手动循环访问。

**处理状态：** 已修复。添加 `getCells()`（返回不可变列表）和 `stream()` 方法。

---

## 总结

| 严重程度 | 数量 | 已修复 | 待处理 | 关键主题 |
|----------|------|--------|--------|----------|
| 中 (Medium) | 5 | 4 | 1 | NumberFormatException、rawRow 越界、XSSF 图片 dimension、内部列表暴露 |
| 低 (Low) | 13 | 9 | 4 | 迭代器异常、可见性、方法命名、反射、残留注释、API 扩展 |

**已修复 13 项 / 18 项**。剩余 5 项为测试覆盖待补充（#13-#16 及其他涉及测试的项）。

与 5.0.0 审查（6 严重 + 5 高危 + 7 中等 + 6 低危）相比，当前代码质量已大幅提升。残留问题以中低风险的设计改进和测试补充为主，无严重或高危问题。
