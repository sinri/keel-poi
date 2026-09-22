# keel-poi 5.0.3 项目审查

审查日期：2026-09-20。仓库：https://github.com/sinri/keel-poi 。
分支：`dev-5.0.3`；HEAD：`d51485f488bdc6516147a93580863e0ee8c1c5cc`。
以当前工作区为准；开始审查前已有 `gradle.properties` 修改和未跟踪的 `README.md`，本次保留这些内容。

## 范围与验证

检查 CSV 读写、Excel 打开/创建/保存、矩阵与模板、迭代器、图片提取、模块声明、Gradle 依赖和现有测试；对照 5.0.2 审查报告，未将已修复问题重新列为发现。
项目定位为 Java 17、Apache POI / excel-streaming-reader 与 Keel / Vert.x Future 的集成库。

- 当前配置 `./gradlew test` 成功，现有 36 项测试无失败；本次再次执行命中 UP-TO-DATE。
- 使用当前 main runtimeClasspath 运行独立 Java 最小复现，确认下述四项问题。复现未修改生产代码或正式测试。
- GitHub 未关闭 Issue 查询结果为空（仅核对未关闭项）。
- 上一轮依赖评估中，临时覆盖 reader 为 5.3.2 后编译及 36 项测试通过；这不代表以下边界问题已解决。
- 未进行大文件压力测试、所有 Excel 格式组合测试或发布流程验证；本报告不声称穷尽所有缺陷。

## 决策记录

| 编号 | 优先级 | 问题 | 状态 | Issue / 提交 |
|---|---|---|---|---|
| R1 | P1 | 默认 XLS 输入流自动识别失败 | 用户决定建立 Issue；已创建，待处理 | [#8](https://github.com/sinri/keel-poi/issues/8) |
| R2 | P1 | 稀疏工作表按表头行索引读取时选错行或丢失数据 | 用户决定建立 Issue；已创建，待处理 | [#9](https://github.com/sinri/keel-poi/issues/9) |
| R3 | P2 | UTF-16 CSV 输出重复插入 BOM | 用户决定建立 Issue；已创建，待处理 | [#10](https://github.com/sinri/keel-poi/issues/10) |
| R4 | P2 | CSV 多行字段中的 CR/CRLF 被静默改为 LF | 用户决定建立 Issue；已创建，待处理 | [#11](https://github.com/sinri/keel-poi/issues/11) |

P1 表示建议优先处理的核心功能错误；P2 表示特定输入下的数据正确性问题。
每项可选择“不处理 / 建 GitHub Issue / 快速修复”。修复后需完整验证，提交需另获用户同意。

## R1：默认 XLS 输入流自动识别失败

处理记录（2026-09-20）：用户选择“建立 Issue”，已创建 [Issue #8](https://github.com/sinri/keel-poi/issues/8)。本项未修改生产代码。

代码：`src/main/java/io/github/sinri/keel/integration/poi/excel/KeelSheets.java:113-128`。

当只配置 `setInputStream(...)`、不配置 `setUseXlsx(...)` 时，代码先复制整个流并尝试 `XSSFWorkbook`，仅捕获 `IOException` 后回退到 `HSSFWorkbook`。
当前 POI 对合法 XLS/OLE2 输入抛出 `OLE2NotOfficeXmlFileException`，该异常不被此分支捕获，导致 Future 失败，无法执行 XLS 回退。

复现：使用 `HSSFWorkbook` 创建带一行数据的合法 XLS，写入字节数组，然后调用：

```java
KeelSheets.useSheets(
    new SheetsOpenOptions().setInputStream(new ByteArrayInputStream(xlsBytes)),
    sheets -> Future.succeededFuture(sheets.getSheetCount()));
```

期望：打开成功，返回 1。实际：失败原因为 `OLE2NotOfficeXmlFileException`。
影响：调用默认流入口读取传统 `.xls` 失败；明确指定 `setUseXlsx(false)` 可绕过。

建议：采用 `WorkbookFactory.create(InputStream)` 统一格式探测，并从实际 Workbook 类型确定 reader 类型；保持明确格式选项的既定语义，明确输入流所有权。该方案也可消除当前自动识别分支额外的 `readAllBytes()` 副本。
验证要求：默认流入口 XLS/XLSX 均成功、显式格式路径正常、非法输入返回失败且生命周期正确。

## R2：稀疏工作表表头行索引错误

处理记录（2026-09-20）：用户选择“建立 Issue”，已创建 [Issue #9](https://github.com/sinri/keel-poi/issues/9)。本项未修改生产代码。

代码：`src/main/java/io/github/sinri/keel/integration/poi/excel/KeelSheet.java:370`、`:423`、`:505`、`:558`。

四个矩阵读取入口使用自增计数器与 `headerRowIndex` 比较，但底层 `sheet.rowIterator()` 仅遍历存在的行。
只要目标表头之前存在未创建的行，计数器就不等于工作表行索引，与参数文档及 `readRow(i)` 的坐标语义不一致。

复现：在 XSSFWorkbook 中只创建索引 2 的表头 `header`、索引 3 的数据 `data`，调用 `readAllRowsToMatrix(2, 1, null)`。
期望：表头为 `[header]`，数据为 `[[data]]`。实际：表头为 `[]`，未读取到对应数据。
同步普通矩阵路径已运行复现；同步模板及两个异步路径经源码确认使用同样计数方式，尚未分别运行复现。

影响：有前导空行或稀疏行的文件可能被静默读为空矩阵、选错表头，模板路径也可能最终触发空值异常。
建议：使用 `row.getRowNum()` 匹配物理行索引，并明确定义目标表头不存在时的失败行为；四个入口统一处理。
验证要求：连续行、前导缺行、中间缺行、缺失表头以及同步/异步入口结果一致。

## R3：UTF-16 CSV 输出重复 BOM

处理记录（2026-09-20）：用户选择“建立 Issue”，已创建 [Issue #10](https://github.com/sinri/keel-poi/issues/10)。本项未修改生产代码。

代码：`src/main/java/io/github/sinri/keel/integration/poi/csv/KeelCsvWriter.java:145-146`，调用方包括 `writeCell`、`writeRowEnding`、`blockWriteRow`。

每个输出片段分别执行 `String.getBytes(charset)`。对于 `StandardCharsets.UTF_16`，每次独立编码均产生 BOM，所以字段、分隔符、行结束符甚至后续行之前都会写入 BOM。

复现：UTF-16 writer 依次执行 `writeCell("a")`、`writeCell("b")`、`writeRowEnding()`、`blockWriteRow(List.of("c", "d"))`。
将完整字节流按 UTF-16 解码，实际为：

```text
a<BOM>,<BOM>b<BOM><CR><LF><BOM>c,d<CR><LF>
```

期望解码内容：`a,b\r\nc,d\r\n`。内部 BOM 会变成字段数据，影响列名匹配及字符串比较。
默认 UTF-8 路径不受此 BOM 问题影响。
建议：整个 writer 生命周期共用 `OutputStreamWriter` 或持续的编码器，统一处理 flush/close。
验证要求：UTF-8、UTF-16 逐单元格/整行/混合写入，回读结果完全一致，BOM 仅出现在文件起始位置。

## R4：CSV 多行字段丢失原始换行形式

处理记录（2026-09-20）：用户选择“建立 Issue”，已创建 [Issue #11](https://github.com/sinri/keel-poi/issues/11)。本项未修改生产代码。

代码：`src/main/java/io/github/sinri/keel/integration/poi/csv/KeelCsvReader.java:104`、`:160-161`。

读取器用 `BufferedReader.readLine()` 去除原始行结束符，再对引号未闭合的字段追加固定 `\n`。
因此合法引号字段内的 `\r\n` 或单独 `\r` 被改成 `\n`。Writer 会保留这些字符，Reader 却不能原样回读。

复现输入：`"a\r\nb",c\r\n`（外层双引号为 CSV 字段引号）。
期望首字段：`a\r\nb`；实际首字段：`a\nb`。该差异已运行确认。

影响：包含多行文本的 CSV 导入/导出无法无损往返，影响精确比较及原文存储。
建议：字符级解析并区分引号内外的 CR、LF、CRLF；若有意归一化，需要公开明确该契约并提供保留原文选项。
验证要求：字段内 CR/LF/CRLF 均原样保留，字段外正确划分记录，同时保持双引号转义及多字符分隔符能力。

## 依赖维护备注（不计入上述缺陷）

可单独将 `excelStreamingReaderVersion` 从 5.2.0 更新为 5.3.2；前一轮已临时验证。建议与功能修复分开记录，避免把依赖更新当作这些应用层问题的解决方案。

## 后续执行规则

按 R1 → R4 逐项取得决策并更新本文件；不自动建立 Issue、修复或提交。依据用户明确调用的 github-project-review 技能要求：“基于文档，by问题一个个与用户沟通……以获得用户的决策（不处理、建GitHub的Issue、快速修复）”。
