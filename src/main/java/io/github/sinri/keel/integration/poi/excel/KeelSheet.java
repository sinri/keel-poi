package io.github.sinri.keel.integration.poi.excel;

import io.github.sinri.keel.base.Keel;
import io.github.sinri.keel.core.utils.value.ValueBox;
import io.github.sinri.keel.integration.poi.excel.entity.*;
import io.vertx.core.Future;
import org.apache.poi.ss.usermodel.*;
import org.jspecify.annotations.NullMarked;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;
import java.util.function.Function;


/**
 * Excel 工作表操作类，提供对 Excel 工作表的读写操作。
 * <p>
 * 该类封装了 Apache POI 的工作表操作，提供了更简洁的 API。
 *
 * @since 5.0.0
 */
@NullMarked
public class KeelSheet {
    private final Sheet sheet;
    /**
     * 公式求值器值盒子，用于存储公式求值器实例。
     * <p>
     * 公式求值器用于处理工作表中的公式单元格。
     *
     */
    private final ValueBox<FormulaEvaluator> formulaEvaluatorBox;


    /**
     * 工作表读取器类型，用于标识工作表的读取方式。
     * 在写入模式下此字段为 null。
     */

    protected @Nullable KeelSheetsReaderType sheetsReaderType;

    /**
     * 使用指定的工作表读取器类型和 POI 工作表实例加载工作表，不使用公式求值器。
     * <p>
     * 即，公式单元格将被解析为字符串形式。
     *
     * @param sheetsReaderType 工作表读取器类型
     * @param sheet            POI 工作表实例
     */
    KeelSheet(@Nullable KeelSheetsReaderType sheetsReaderType, Sheet sheet) {
        this(sheetsReaderType, sheet, new ValueBox<>());
    }

    /**
     * 使用指定的公式求值器加载工作表，支持三种类型的单元格公式求值器：无、缓存和求值。
     *
     * @param sheetsReaderType    工作表读取器类型
     * @param sheet               POI 工作表实例
     * @param formulaEvaluatorBox 公式求值器值盒子
     */
    KeelSheet(@Nullable KeelSheetsReaderType sheetsReaderType, Sheet sheet, ValueBox<FormulaEvaluator> formulaEvaluatorBox) {
        this.sheetsReaderType = sheetsReaderType;
        this.sheet = sheet;
        this.formulaEvaluatorBox = formulaEvaluatorBox;
    }

    /**
     * 自动检测一行中非空单元格的数量。
     * 该方法从索引零开始计算到最后一个非零单元格的数量。如果没有单元格，则返回 0。
     *
     * @param row 包含单元格的 POI 行
     * @return 从索引零到最后一个非零单元格的单元格数量
     */
    private static int autoDetectNonBlankColumnCountInOneRow(Row row) {
        short firstCellNum = row.getFirstCellNum();
        if (firstCellNum < 0) {
            return 0;
        }
        int i;
        for (i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                break;
            }
            if (cell.getCellType() != CellType.NUMERIC) {
                String stringCellValue = cell.getStringCellValue();
                if (stringCellValue == null || stringCellValue.isBlank()) break;
            }
        }
        return i;
    }

    /**
     * 将单元格内容转换为字符串。
     *
     * @param cell                单元格（可能为 null）
     * @param formulaEvaluatorBox 公式求值器值盒子
     * @return 单元格内容的字符串表示
     */
    public static String dumpCellToString(
            @Nullable Cell cell,
            ValueBox<FormulaEvaluator> formulaEvaluatorBox
    ) {
        if (cell == null) return "";
        CellType cellType = cell.getCellType();
        String s;
        if (cellType == CellType.NUMERIC) {
            double numericCellValue = cell.getNumericCellValue();
            s = String.valueOf(numericCellValue);
        } else if (cellType == CellType.FORMULA) {
            if (formulaEvaluatorBox.isValueAlreadySet()) {
                CellType formulaResultType;

                FormulaEvaluator formulaEvaluator = formulaEvaluatorBox.getValue();

                if (formulaEvaluator == null) {
                    formulaResultType = cell.getCachedFormulaResultType();
                } else {
                    formulaResultType = formulaEvaluator.evaluateFormulaCell(cell);
                }
                s = switch (formulaResultType) {
                    case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                    case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                    case STRING -> String.valueOf(cell.getStringCellValue());
                    case ERROR -> String.valueOf(cell.getErrorCellValue());
                    default -> throw new RuntimeException("FormulaResultType unknown");
                };
            } else {
                return cell.getStringCellValue();
            }
        } else {
            s = cell.getStringCellValue();
        }
        return Objects.requireNonNull(s);
    }

    /**
     * 将行数据转换为原始行列表。
     *
     * @param row                 POI 行对象
     * @param maxColumns          最大列数
     * @param sheetRowFilter      工作表行过滤器（可选）
     * @param formulaEvaluatorBox 公式求值器值盒子
     * @return 原始行数据列表，如果行被过滤器丢弃则返回 null
     */
    public static @Nullable List<String> dumpRowToRawRow(
            Row row,
            int maxColumns,
            @Nullable SheetRowFilter sheetRowFilter,
            ValueBox<FormulaEvaluator> formulaEvaluatorBox
    ) {
        List<String> rowDatum = new ArrayList<>();

        for (int i = 0; i < maxColumns; i++) {
            Cell cell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            String s = dumpCellToString(cell, formulaEvaluatorBox);
            rowDatum.add(s);
        }

        if (sheetRowFilter != null) {
            if (sheetRowFilter.shouldThrowThisRawRow(rowDatum)) {
                return null;
            }
        }

        return rowDatum;
    }

    /**
     * 获取工作表读取器类型。
     *
     * @return 工作表读取器类型，可能为 null
     */

    public @Nullable KeelSheetsReaderType getSheetsReaderType() {
        return sheetsReaderType;
    }

    /**
     * 设置工作表读取器类型。
     *
     * @param sheetsReaderType 工作表读取器类型
     * @return 当前工作表对象，支持链式调用
     */
    public KeelSheet setSheetsReaderType(@Nullable KeelSheetsReaderType sheetsReaderType) {
        this.sheetsReaderType = sheetsReaderType;
        return this;
    }

    /**
     * 获取工作表绘制对象，用于处理工作表中的图形元素。
     *
     * @return 工作表绘制对象
     */
    private KeelSheetDrawing getDrawing() {
        return new KeelSheetDrawing(this);
    }

    /**
     * 获取从工作表中读取的图片列表。
     *
     * @return 从工作表中读取的 {@link KeelPictureInSheet} 图片对象列表
     */
    public List<KeelPictureInSheet> getPictures() {
        return getDrawing().getPictures();
    }

    /**
     * 获取工作表的名称。
     *
     * @return 工作表的名称
     */
    public String getName() {
        return sheet.getSheetName();
    }

    /**
     * 读取指定行索引的行。
     *
     * @param i 行索引
     * @return 指定行索引的 POI 行对象
     */
    public Row readRow(int i) {
        return sheet.getRow(i);
    }

    /**
     * 获取行迭代器。
     *
     * @return 行迭代器，用于遍历工作表中的所有行
     */
    public Iterator<Row> getRowIterator() {
        return sheet.rowIterator();
    }

    /**
     * 读取指定行索引的原始行数据。
     *
     * @param i              行索引
     * @param maxColumns     最大列数
     * @param sheetRowFilter 工作表行过滤器（可选）
     * @return 原始行数据列表
     */
    public @Nullable List<String> readRawRow(int i, int maxColumns, @Nullable SheetRowFilter sheetRowFilter) {
        var row = readRow(i);
        return dumpRowToRawRow(row, maxColumns, sheetRowFilter, this.formulaEvaluatorBox);
    }

    /**
     * 获取原始行迭代器。
     *
     * @param maxColumns     最大列数
     * @param sheetRowFilter 工作表行过滤器（可选）
     * @return 原始行数据列表的迭代器
     */
    public Iterator<List<String>> getRawRowIterator(int maxColumns, @Nullable SheetRowFilter sheetRowFilter) {
        Iterator<Row> rowIterator = getRowIterator();
        return new Iterator<>() {
            @Override
            public boolean hasNext() {
                return rowIterator.hasNext();
            }

            @Override
            public @Nullable List<String> next() {
                Row row = rowIterator.next();
                return dumpRowToRawRow(row, maxColumns, sheetRowFilter, formulaEvaluatorBox);
            }
        };
    }

    /**
     * 以阻塞方式读取所有行，并对每一行执行指定的操作。
     *
     * @param rowConsumer 行消费者，用于处理每一行数据
     */
    public final void readAllRows(Consumer<Row> rowConsumer) {
        Iterator<Row> it = getRowIterator();

        while (it.hasNext()) {
            Row row = it.next();
            rowConsumer.accept(row);
        }
    }

    /**
     * 获取原始的 Apache POI 工作表实例。
     *
     * @return 原始的 Apache POI 工作表实例
     */
    public Sheet getSheet() {
        return sheet;
    }

    /**
     * 以阻塞方式读取所有行并转换为矩阵，遵循以下规则：(1) 第一行作为表头，(2) 自动检测列数，(3) 抛弃空行。
     *
     * @return 读取的矩阵对象
     */
    public final KeelSheetMatrix readAllRowsToMatrix() {
        return readAllRowsToMatrix(0, 0, SheetRowFilter.toThrowEmptyRows());
    }

    /**
     * 以阻塞方式读取所有行并转换为矩阵，表头行之前的行将被丢弃！
     *
     * @param headerRowIndex 表头行索引，0 表示第一行，依此类推
     * @param maxColumns     预设列数，如果需要自动检测则为零或负数
     * @param sheetRowFilter 工作表行过滤器（可选）
     * @return 读取的矩阵对象
     */
    public final KeelSheetMatrix readAllRowsToMatrix(int headerRowIndex, int maxColumns, @Nullable SheetRowFilter sheetRowFilter) {
        if (headerRowIndex < 0) throw new IllegalArgumentException("headerRowIndex less than zero");

        KeelSheetMatrix keelSheetMatrix = new KeelSheetMatrix();
        AtomicInteger rowIndex = new AtomicInteger(0);

        AtomicInteger checkColumnsRef = new AtomicInteger();
        if (maxColumns > 0) {
            checkColumnsRef.set(maxColumns);
        }

        readAllRows(row -> {
            int currentRowIndex = rowIndex.get();
            if (headerRowIndex == currentRowIndex) {
                if (checkColumnsRef.get() == 0) {
                    checkColumnsRef.set(autoDetectNonBlankColumnCountInOneRow(row));
                }
                List<String> headerRow = dumpRowToRawRow(row, checkColumnsRef.get(), sheetRowFilter, formulaEvaluatorBox);
                if (headerRow == null) {
                    throw new NullPointerException("Header Row is not valid");
                }
                keelSheetMatrix.setHeaderRow(headerRow);
            } else if (headerRowIndex < currentRowIndex) {
                var x = dumpRowToRawRow(row, checkColumnsRef.get(), sheetRowFilter, formulaEvaluatorBox);
                if (x != null) {
                    keelSheetMatrix.addRow(x);
                }
            }

            rowIndex.incrementAndGet();
        });

        return keelSheetMatrix;
    }

    /**
     * 以阻塞方式读取所有行并转换为模板化矩阵，遵循以下规则：(1) 第一行作为表头，(2) 自动检测列数，(3) 抛弃空行。
     *
     * @return 读取的模板化矩阵对象
     */
    public final KeelSheetTemplatedMatrix readAllRowsToTemplatedMatrix() {
        return readAllRowsToTemplatedMatrix(0, 0, SheetRowFilter.toThrowEmptyRows());
    }

    /**
     * 以阻塞方式读取所有行并转换为模板化矩阵，表头行之前的行将被丢弃！
     *
     * @param headerRowIndex 表头行索引，0 表示第一行，依此类推
     * @param maxColumns     预设列数，如果需要自动检测则为零或负数
     * @param sheetRowFilter 工作表行过滤器（可选）
     * @return 读取的模板化矩阵对象
     */
    public final KeelSheetTemplatedMatrix readAllRowsToTemplatedMatrix(int headerRowIndex, int maxColumns, @Nullable SheetRowFilter sheetRowFilter) {
        if (headerRowIndex < 0) throw new IllegalArgumentException("headerRowIndex less than zero");

        AtomicInteger checkColumnsRef = new AtomicInteger();
        if (maxColumns > 0) {
            checkColumnsRef.set(maxColumns);
        }

        AtomicInteger rowIndex = new AtomicInteger(0);
        AtomicReference<@Nullable KeelSheetTemplatedMatrix> templatedMatrixRef = new AtomicReference<>();


        readAllRows(row -> {
            int currentRowIndex = rowIndex.get();
            if (currentRowIndex == headerRowIndex) {
                if (checkColumnsRef.get() == 0) {
                    checkColumnsRef.set(autoDetectNonBlankColumnCountInOneRow(row));
                }

                var rowDatum = dumpRowToRawRow(row, checkColumnsRef.get(), sheetRowFilter, formulaEvaluatorBox);
                if (rowDatum == null) throw new NullPointerException("Header Row is not valid");
                KeelSheetMatrixRowTemplate rowTemplate = KeelSheetMatrixRowTemplate.create(rowDatum);
                KeelSheetTemplatedMatrix templatedMatrix = KeelSheetTemplatedMatrix.create(rowTemplate);
                templatedMatrixRef.set(templatedMatrix);
            } else if (currentRowIndex > headerRowIndex) {
                var rowDatum = dumpRowToRawRow(row, checkColumnsRef.get(), sheetRowFilter, formulaEvaluatorBox);
                if (rowDatum != null) {
                    var r = templatedMatrixRef.get();
                    Objects.requireNonNull(r).addRawRow(rowDatum);
                }
            }
            rowIndex.incrementAndGet();
        });
        var r = templatedMatrixRef.get();
        return Objects.requireNonNull(r);
    }

    /**
     * 异步读取所有行，并对每一行执行指定的操作。
     * 建议在工作线程上下文中调用此方法。
     * 逐行处理效率不够高。
     *
     * @param rowFunc 行处理函数，用于处理每一行数据
     * @return 表示操作完成的 Future
     */
    public final Future<Void> readAllRowsAsync(Keel keel, Function<Row, Future<Void>> rowFunc) {
        return keel.asyncCallIteratively(getRowIterator(), rowFunc);
    }

    /**
     * 异步读取所有行，并按批次对行执行指定的操作。
     * 建议在工作线程上下文中调用此方法。
     *
     * @param rowsFunc  行批处理函数，用于处理行批次数据
     * @param batchSize 批次大小
     * @return 表示操作完成的 Future
     */
    public final Future<Void> readAllRowsAsync(Keel keel, Function<List<Row>, Future<Void>> rowsFunc, int batchSize) {
        return keel.asyncCallIteratively(
                getRowIterator(),
                rowsFunc,
                batchSize
        );
    }

    /**
     * 异步读取所有行并转换为矩阵，遵循以下规则：(1) 第一行作为表头，(2) 自动检测列数，(3) 抛弃空行。
     *
     * @return 表示矩阵读取完成的 Future
     */
    public final Future<KeelSheetMatrix> readAllRowsToMatrixAsync(Keel keel) {
        return readAllRowsToMatrixAsync(keel, 0, 0, SheetRowFilter.toThrowEmptyRows());
    }

    /**
     * 异步读取所有行并转换为矩阵，表头行之前的行将被丢弃！
     *
     * @param headerRowIndex 表头行索引，0 表示第一行，依此类推
     * @param maxColumns     预设列数，如果需要自动检测则为零或负数
     * @param sheetRowFilter 工作表行过滤器（可选）
     * @return 表示矩阵读取完成的 Future
     */
    public final Future<KeelSheetMatrix> readAllRowsToMatrixAsync(Keel keel, int headerRowIndex, int maxColumns, @Nullable SheetRowFilter sheetRowFilter) {
        if (headerRowIndex < 0) throw new IllegalArgumentException("headerRowIndex less than zero");

        AtomicInteger checkColumnsRef = new AtomicInteger();
        if (maxColumns > 0) {
            checkColumnsRef.set(maxColumns);
        }

        KeelSheetMatrix keelSheetMatrix = new KeelSheetMatrix();
        AtomicInteger rowIndex = new AtomicInteger(0);

        return readAllRowsAsync(keel, rows -> {
            rows.forEach(row -> {
                int currentRowIndex = rowIndex.get();
                if (headerRowIndex == currentRowIndex) {
                    if (checkColumnsRef.get() == 0) {
                        checkColumnsRef.set(autoDetectNonBlankColumnCountInOneRow(row));
                    }
                    var headerRow = dumpRowToRawRow(row, checkColumnsRef.get(), sheetRowFilter, formulaEvaluatorBox);
                    if (headerRow == null) {
                        throw new NullPointerException("Header Row is not valid");
                    }
                    keelSheetMatrix.setHeaderRow(headerRow);
                } else if (headerRowIndex < currentRowIndex) {
                    List<String> rawRow = dumpRowToRawRow(row, checkColumnsRef.get(), sheetRowFilter, formulaEvaluatorBox);
                    if (rawRow != null) {
                        keelSheetMatrix.addRow(rawRow);
                    }
                }
                rowIndex.incrementAndGet();
            });
            return Future.succeededFuture();
        }, 1000)
                .compose(v -> Future.succeededFuture(keelSheetMatrix));
    }

    /**
     * 异步读取所有行并转换为模板化矩阵，遵循以下规则：(1) 第一行作为表头，(2) 自动检测列数，(3) 抛弃空行。
     *
     * @return 表示模板化矩阵读取完成的 Future
     */
    public final Future<KeelSheetTemplatedMatrix> readAllRowsToTemplatedMatrixAsync(Keel keel) {
        return readAllRowsToTemplatedMatrixAsync(keel, 0, 0, SheetRowFilter.toThrowEmptyRows());
    }

    /**
     * 异步读取所有行并转换为模板化矩阵，表头行之前的行将被丢弃！
     *
     * @param headerRowIndex 表头行索引，0 表示第一行，依此类推
     * @param maxColumns     预设列数，如果需要自动检测则为零或负数
     * @param sheetRowFilter 工作表行过滤器（可选）
     * @return 表示模板化矩阵读取完成的 Future
     */
    public final Future<KeelSheetTemplatedMatrix> readAllRowsToTemplatedMatrixAsync(Keel keel, int headerRowIndex, int maxColumns, @Nullable SheetRowFilter sheetRowFilter) {
        if (headerRowIndex < 0) throw new IllegalArgumentException("headerRowIndex less than zero");

        AtomicInteger checkColumnsRef = new AtomicInteger();
        if (maxColumns > 0) {
            checkColumnsRef.set(maxColumns);
        }

        AtomicInteger rowIndex = new AtomicInteger(0);
        AtomicReference<@Nullable KeelSheetTemplatedMatrix> templatedMatrixRef = new AtomicReference<>();

        return readAllRowsAsync(keel, rows -> {
            rows.forEach(row -> {
                int currentRowIndex = rowIndex.get();
                if (currentRowIndex == headerRowIndex) {
                    if (checkColumnsRef.get() == 0) {
                        checkColumnsRef.set(autoDetectNonBlankColumnCountInOneRow(row));
                    }

                    var rowDatum = dumpRowToRawRow(row, checkColumnsRef.get(), sheetRowFilter, formulaEvaluatorBox);
                    if (rowDatum == null) {
                        throw new NullPointerException("Header Row is not valid");
                    }
                    KeelSheetMatrixRowTemplate rowTemplate = KeelSheetMatrixRowTemplate.create(rowDatum);
                    KeelSheetTemplatedMatrix templatedMatrix = KeelSheetTemplatedMatrix.create(rowTemplate);
                    templatedMatrixRef.set(templatedMatrix);
                } else if (currentRowIndex > headerRowIndex) {
                    List<String> rowDatum = dumpRowToRawRow(row, checkColumnsRef.get(), sheetRowFilter, formulaEvaluatorBox);
                    if (rowDatum != null) {
                        KeelSheetTemplatedMatrix matrix = templatedMatrixRef.get();
                        Objects.requireNonNull(matrix).addRawRow(rowDatum);
                    }
                }
                rowIndex.incrementAndGet();
            });
            return Future.succeededFuture();
        }, 1000)
                .compose(v -> {
                    KeelSheetTemplatedMatrix r = templatedMatrixRef.get();
                    return Future.succeededFuture(Objects.requireNonNull(r));
                });
    }


    /**
     * 以阻塞方式将行数据写入工作表，从指定的行索引和单元格索引开始。
     *
     * @param rowData        行数据列表
     * @param sinceRowIndex  起始行索引
     * @param sinceCellIndex 起始单元格索引
     */
    public void writeAllRows(List<List<String>> rowData, int sinceRowIndex, int sinceCellIndex) {
        for (int rowIndex = 0; rowIndex < rowData.size(); rowIndex++) {
            Row row = sheet.getRow(sinceRowIndex + rowIndex);
            if (row == null) {
                row = sheet.createRow(sinceRowIndex + rowIndex);
            }
            var rowDatum = rowData.get(rowIndex);
            writeToRow(row, rowDatum, sinceCellIndex);
        }
    }

    /**
     * 以阻塞方式将行数据写入工作表，从第 0 行第 0 列开始。
     *
     * @param rowData 行数据列表
     */
    public void writeAllRows(List<List<String>> rowData) {
        writeAllRows(rowData, 0, 0);
    }

    /**
     * 以阻塞方式将矩阵数据写入工作表。
     * 如果矩阵没有表头行，则直接写入原始行数据；否则先写入表头行，再写入原始行数据。
     *
     * @param matrix 矩阵数据
     */
    public void writeMatrix(KeelSheetMatrix matrix) {
        if (matrix.getHeaderRow().isEmpty()) {
            writeAllRows(matrix.getRawRowList(), 0, 0);
        } else {
            writeAllRows(List.of(matrix.getHeaderRow()), 0, 0);
            writeAllRows(matrix.getRawRowList(), 1, 0);
        }
    }

    /**
     * 异步将矩阵数据写入工作表。
     * 如果矩阵有表头行，则先写入表头行，再写入原始行数据。
     *
     * @param matrix 矩阵数据
     * @return 表示写入操作完成的 Future
     */
    public Future<Void> writeMatrixAsync(Keel keel, KeelSheetMatrix matrix) {
        AtomicInteger rowIndexRef = new AtomicInteger(0);
        if (!matrix.getHeaderRow().isEmpty()) {
            writeAllRows(List.of(matrix.getHeaderRow()), 0, 0);
            rowIndexRef.incrementAndGet();
        }

        return keel.asyncCallIteratively(matrix.getRawRowList().iterator(), rawRows -> {
            writeAllRows(matrix.getRawRowList(), rowIndexRef.get(), 0);
            rowIndexRef.addAndGet(rawRows.size());
            return Future.succeededFuture();
        }, 1000);
    }

    /**
     * 以阻塞方式将模板化矩阵数据写入工作表。
     * 首先写入模板的列名称作为表头，然后逐行写入模板化行数据。
     *
     * @param templatedMatrix 模板化矩阵数据
     */
    public void writeTemplatedMatrix(KeelSheetTemplatedMatrix templatedMatrix) {
        AtomicInteger rowIndexRef = new AtomicInteger(0);
        writeAllRows(List.of(templatedMatrix.getTemplate().getColumnNames()), 0, 0);
        rowIndexRef.incrementAndGet();
        templatedMatrix.getRows()
                       .forEach(templatedRow -> writeAllRows(List.of(templatedRow.getRawRow()), rowIndexRef.get(), 0));
    }

    /**
     * 异步将模板化矩阵数据写入工作表。
     * 首先写入模板的列名称作为表头，然后逐行写入模板化行数据。
     *
     * @param templatedMatrix 模板化矩阵数据
     * @return 表示写入操作完成的 Future
     */
    public Future<Void> writeTemplatedMatrixAsync(Keel keel, KeelSheetTemplatedMatrix templatedMatrix) {
        AtomicInteger rowIndexRef = new AtomicInteger(0);
        writeAllRows(List.of(templatedMatrix.getTemplate().getColumnNames()), 0, 0);
        rowIndexRef.incrementAndGet();

        return keel.asyncCallIteratively(templatedMatrix.getRawRows().iterator(), rawRows -> {
            writeAllRows(rawRows, rowIndexRef.get(), 0);
            rowIndexRef.addAndGet(rawRows.size());
            return Future.succeededFuture();
        }, 1000);
    }


    private void writeToRow(Row row, List<String> rowDatum, int sinceCellIndex) {
        for (int cellIndex = 0; cellIndex < rowDatum.size(); cellIndex++) {
            var cellDatum = rowDatum.get(cellIndex);

            Cell cell = row.getCell(cellIndex + sinceCellIndex);
            if (cell == null) {
                cell = row.createCell(cellIndex + sinceCellIndex);
            }

            cell.setCellValue(cellDatum);
        }
    }

    /**
     * 获取矩阵行迭代器。
     * 该迭代器将原始行数据转换为矩阵行对象，便于按行处理矩阵数据。
     *
     * @param maxColumns     最大列数
     * @param sheetRowFilter 工作表行过滤器（可选）
     * @return 矩阵行迭代器
     */
    public Iterator<KeelSheetMatrixRow> getMatrixRowIterator(int maxColumns, @Nullable SheetRowFilter sheetRowFilter) {
        Iterator<List<String>> rawRowIterator = this.getRawRowIterator(maxColumns, sheetRowFilter);
        return new Iterator<>() {
            @Override
            public boolean hasNext() {
                return rawRowIterator.hasNext();
            }

            @Override
            public KeelSheetMatrixRow next() {
                return new KeelSheetMatrixRow(rawRowIterator.next());
            }
        };

    }

    /**
     * 获取模板化矩阵行迭代器。
     * 该迭代器将原始行数据转换为模板化矩阵行对象，便于按行处理模板化矩阵数据。
     *
     * @param template       矩阵行模板
     * @param maxColumns     最大列数
     * @param sheetRowFilter 工作表行过滤器（可选）
     * @return 模板化矩阵行迭代器
     */
    public Iterator<KeelSheetMatrixTemplatedRow> getTemplatedMatrixRowIterator(
            KeelSheetMatrixRowTemplate template,
            int maxColumns,
            @Nullable SheetRowFilter sheetRowFilter
    ) {
        Iterator<List<String>> rawRowIterator = this.getRawRowIterator(maxColumns, sheetRowFilter);
        return new Iterator<>() {
            @Override
            public boolean hasNext() {
                return rawRowIterator.hasNext();
            }

            @Override
            public KeelSheetMatrixTemplatedRow next() {
                return KeelSheetMatrixTemplatedRow.create(template, rawRowIterator.next());
            }
        };
    }

    /**
     * 获取与工作表关联的公式求值器。
     *
     * @return 如果可用则返回 {@link FormulaEvaluator} 实例；否则返回 null。
     */
    @Nullable
    public FormulaEvaluator getFormulaEvaluator() {
        return formulaEvaluatorBox.getValue();
    }
}
