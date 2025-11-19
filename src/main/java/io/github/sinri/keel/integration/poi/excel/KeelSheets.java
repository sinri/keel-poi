package io.github.sinri.keel.integration.poi.excel;

import com.github.pjfanning.xlsx.impl.StreamingWorkbook;
import io.github.sinri.keel.core.utils.value.ValueBox;
import io.vertx.core.Closeable;
import io.vertx.core.Completable;
import io.vertx.core.Future;
import io.vertx.core.Promise;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.io.*;
import java.util.Objects;
import java.util.function.Function;

/**
 * Excel 工作簿管理类，提供对 Excel 工作簿的创建、打开、保存等操作。
 * <p>
 * 该类封装了 Apache POI 的工作簿操作，提供了更简洁的 API。
 *
 * @since 5.0.0
 */
public class KeelSheets implements Closeable {
    /**
     * 公式求值器，用于计算 Excel 公式单元格的值。
     * 该字段在需要公式求值时初始化。
     *
     * @since 5.0.0
     */
    private final @Nullable FormulaEvaluator formulaEvaluator;
    protected @NotNull Workbook autoWorkbook;
    /**
     * This field is null for write mode.
     */
    @Nullable
    protected KeelSheetsReaderType sheetsReaderType;

    /**
     * 受保护的构造函数，使用指定的工作表读取器类型和工作簿实例创建工作簿实例。
     *
     * @param sheetsReaderType 工作表读取器类型
     * @param workbook         工作簿实例
     * @since 5.0.0
     */
    protected KeelSheets(@Nullable KeelSheetsReaderType sheetsReaderType, @NotNull Workbook workbook) {
        this(sheetsReaderType, workbook, false);
    }

    /**
     * 受保护的构造函数，创建一个新的工作簿实例。
     *
     */
    protected KeelSheets() {
        this(null, null, false);
    }

    /**
     * 受保护的构造函数，使用指定的工作表读取器类型、工作簿实例和公式求值器选项创建工作簿实例。
     *
     * @param sheetsReaderType     工作表读取器类型
     * @param workbook             工作簿实例，如果为 null 则创建新的工作簿
     * @param withFormulaEvaluator 是否启用公式求值器
     */
    private KeelSheets(@Nullable KeelSheetsReaderType sheetsReaderType, @Nullable Workbook workbook, boolean withFormulaEvaluator) {
        this.sheetsReaderType = sheetsReaderType;
        autoWorkbook = Objects.requireNonNullElseGet(workbook, XSSFWorkbook::new);
        if (withFormulaEvaluator) {
            formulaEvaluator = autoWorkbook.getCreationHelper().createFormulaEvaluator();
        } else {
            formulaEvaluator = null;
        }
    }

    /**
     * 使用指定的打开选项打开 Excel 工作簿，并在使用完成后自动关闭。
     * 该方法会自动管理工作簿的生命周期，确保在操作完成后关闭工作簿。
     *
     * @param sheetsOpenOptions 打开工作簿的选项
     * @param usage             使用工作簿的函数
     * @return 表示操作完成的 Future
     */
    public static <T> Future<T> useSheets(@NotNull SheetsOpenOptions sheetsOpenOptions,
                                          @NotNull Function<KeelSheets, Future<T>> usage) {
        return Future.succeededFuture()
                     .compose(v -> {
                         try {
                             KeelSheets keelSheets;
                             if (sheetsOpenOptions.isUseHugeXlsxStreamReading()) {
                                 if (sheetsOpenOptions.getInputStream() != null) {
                                     keelSheets = new KeelSheets(
                                             KeelSheetsReaderType.XLSX_STREAMING,
                                             sheetsOpenOptions.getHugeXlsxStreamingReaderBuilder()
                                                              .open(sheetsOpenOptions.getInputStream())
                                     );
                                 } else if (sheetsOpenOptions.getFile() != null) {
                                     keelSheets = new KeelSheets(
                                             KeelSheetsReaderType.XLSX_STREAMING,
                                             sheetsOpenOptions.getHugeXlsxStreamingReaderBuilder()
                                                              .open(sheetsOpenOptions.getFile())
                                     );
                                 } else {
                                     throw new IOException("No input source!");
                                 }
                             } else {
                                 InputStream inputStream = sheetsOpenOptions.getInputStream();
                                 if (inputStream != null) {
                                     Workbook workbook;
                                     Boolean useXlsx = sheetsOpenOptions.isUseXlsx();
                                     if (useXlsx == null) {
                                         try {
                                             workbook = new XSSFWorkbook(inputStream);
                                             useXlsx = true;
                                         } catch (IOException e) {
                                             try {
                                                 workbook = new HSSFWorkbook(inputStream);
                                                 useXlsx = false;
                                             } catch (IOException ex) {
                                                 throw new RuntimeException(ex);
                                             }
                                         }
                                     } else {
                                         if (useXlsx) {
                                             workbook = new XSSFWorkbook(inputStream);
                                         } else {
                                             workbook = new HSSFWorkbook(inputStream);
                                         }
                                     }
                                     keelSheets = new KeelSheets(
                                             (useXlsx ? KeelSheetsReaderType.XLSX : KeelSheetsReaderType.XLS),
                                             workbook,
                                             sheetsOpenOptions.isWithFormulaEvaluator()
                                     );
                                 } else if (sheetsOpenOptions.getFile() != null) {
                                     Workbook workbook = WorkbookFactory.create(sheetsOpenOptions.getFile());
                                     KeelSheetsReaderType sheetsReaderType1;
                                     if (workbook instanceof XSSFWorkbook || workbook instanceof SXSSFWorkbook) {
                                         sheetsReaderType1 = KeelSheetsReaderType.XLSX;
                                     } else if (workbook instanceof HSSFWorkbook) {
                                         sheetsReaderType1 = KeelSheetsReaderType.XLS;
                                     } else if (workbook instanceof StreamingWorkbook) {
                                         sheetsReaderType1 = KeelSheetsReaderType.XLSX_STREAMING;
                                     } else {
                                         throw new UnsupportedOperationException(
                                                 "unsupported workbook type:" + workbook.getClass().getName()
                                         );
                                     }
                                     keelSheets = new KeelSheets(
                                             sheetsReaderType1,
                                             workbook,
                                             sheetsOpenOptions.isWithFormulaEvaluator()
                                     );
                                 } else {
                                     throw new IOException("No input source!!");
                                 }
                             }
                             return usage.apply(keelSheets)
                                         .andThen(ar -> {
                                             Promise<Void> promise = Promise.promise();
                                             keelSheets.close(promise);
                                         });
                         } catch (IOException e) {
                             return Future.failedFuture(e);
                         }
                     });
    }

    /**
     * 使用指定的创建选项创建 Excel 工作簿，并在使用完成后自动关闭。
     * 该方法会自动管理工作簿的生命周期，确保在操作完成后关闭工作簿。
     *
     * @param sheetsCreateOptions 创建工作簿的选项
     * @param usage               使用工作簿的函数
     * @return 表示操作完成的 Future
     */
    public static <T> Future<T> useSheets(@NotNull SheetsCreateOptions sheetsCreateOptions,
                                          @NotNull Function<KeelSheets, Future<T>> usage) {
        return Future.succeededFuture()
                     .compose(v -> {
                         KeelSheets keelSheets;
                         if (sheetsCreateOptions.isUseXlsx()) {
                             keelSheets = new KeelSheets(null, new XSSFWorkbook(), sheetsCreateOptions.isWithFormulaEvaluator());
                             if (sheetsCreateOptions.isUseStreamWriting()) {
                                 if (keelSheets.autoWorkbook instanceof XSSFWorkbook) {
                                     keelSheets.autoWorkbook = new SXSSFWorkbook((XSSFWorkbook) (keelSheets.autoWorkbook));
                                 } else {
                                     throw new IllegalStateException("Now autoWorkbook is not an instance of XSSFWorkbook.");
                                 }
                             }
                         } else {
                             keelSheets = new KeelSheets(null, new HSSFWorkbook(), sheetsCreateOptions.isWithFormulaEvaluator());
                         }

                         return usage.apply(keelSheets)
                                     .andThen(ar -> {
                                         Promise<Void> promise = Promise.promise();
                                         keelSheets.close(promise);
                                     });
                     });
    }

    /**
     * 根据工作表名称生成工作表读取器，默认解析公式单元格为值。
     *
     * @param sheetName 工作表名称
     * @return 工作表读取器
     */
    public KeelSheet generateReaderForSheet(@NotNull String sheetName) {
        return this.generateReaderForSheet(sheetName, true);
    }

    /**
     * 根据工作表名称生成工作表读取器，可选择是否解析公式单元格为值。
     *
     * @param sheetName               工作表名称
     * @param parseFormulaCellToValue 是否将公式单元格解析为值
     * @return 工作表读取器
     */
    public KeelSheet generateReaderForSheet(@NotNull String sheetName, boolean parseFormulaCellToValue) {
        var sheet = this.getWorkbook().getSheet(sheetName);
        ValueBox<FormulaEvaluator> formulaEvaluatorValueBox = new ValueBox<>();
        if (parseFormulaCellToValue) {
            formulaEvaluatorValueBox.setValue(this.formulaEvaluator);
        }
        return new KeelSheet(sheetsReaderType, sheet, formulaEvaluatorValueBox);
    }

    /**
     * 根据工作表索引生成工作表读取器，默认解析公式单元格为值。
     *
     * @param sheetIndex 工作表索引
     * @return 工作表读取器
     */
    public KeelSheet generateReaderForSheet(int sheetIndex) {
        return this.generateReaderForSheet(sheetIndex, true);
    }

    /**
     * 根据工作表索引生成工作表读取器，可选择是否解析公式单元格为值。
     *
     * @param sheetIndex              工作表索引
     * @param parseFormulaCellToValue 是否将公式单元格解析为值
     * @return 工作表读取器
     */
    public KeelSheet generateReaderForSheet(int sheetIndex, boolean parseFormulaCellToValue) {
        var sheet = this.getWorkbook().getSheetAt(sheetIndex);
        ValueBox<FormulaEvaluator> formulaEvaluatorValueBox = new ValueBox<>();
        if (parseFormulaCellToValue) {
            formulaEvaluatorValueBox.setValue(this.formulaEvaluator);
        }
        return new KeelSheet(sheetsReaderType, sheet, formulaEvaluatorValueBox);
    }

    /**
     * 根据工作表名称生成工作表写入器，并可指定工作表位置。
     *
     * @param sheetName 工作表名称
     * @param pos       工作表位置，如果为 null 则使用默认位置
     * @return 工作表写入器
     */
    public KeelSheet generateWriterForSheet(@NotNull String sheetName, Integer pos) {
        Sheet sheet = this.getWorkbook().createSheet(sheetName);
        if (pos != null) {
            this.getWorkbook().setSheetOrder(sheetName, pos);
        }
        return new KeelSheet(null, sheet, new ValueBox<>(this.formulaEvaluator));
    }

    /**
     * 根据工作表名称生成工作表写入器。
     *
     * @param sheetName 工作表名称
     * @return 工作表写入器
     */
    public KeelSheet generateWriterForSheet(@NotNull String sheetName) {
        return generateWriterForSheet(sheetName, null);
    }

    /**
     * 获取工作簿中的工作表数量。
     *
     * @return 工作簿中的工作表数量
     */
    public int getSheetCount() {
        return autoWorkbook.getNumberOfSheets();
    }

    /**
     * 获取原始的 Apache POI 工作簿实例。
     *
     * @return 原始的 Apache POI 工作簿实例
     */
    @NotNull
    public Workbook getWorkbook() {
        return autoWorkbook;
    }

    /**
     * 将工作簿保存到指定的输出流。
     *
     * @param outputStream 输出流
     */
    public void save(OutputStream outputStream) {
        try {
            autoWorkbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 将工作簿保存到指定的文件。
     *
     * @param file 文件
     */
    public void save(File file) {
        try {
            save(new FileOutputStream(file));
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 将工作簿保存到指定的文件名。
     *
     * @param fileName 文件名
     */
    public void save(String fileName) {
        save(new File(fileName));
    }

    /**
     * 关闭工作簿，释放相关资源。
     *
     * @param completable 关闭操作完成时的通知对象
     */
    @Override
    public void close(Completable<Void> completable) {
        try {
            autoWorkbook.close();
            completable.succeed();
        } catch (IOException e) {
            completable.fail(e);
        }
    }

}
