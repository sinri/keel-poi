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
 * @since 3.0.13
 * @since 3.0.18 Finished Technical Preview.
 * @since 4.0.2 implements `io.vertx.core.Closeable`, remove `java.lang.AutoCloseable`.
 */
public class KeelSheets implements Closeable {
    /**
     * @since 3.1.3
     */
    private final @Nullable FormulaEvaluator formulaEvaluator;
    protected @NotNull Workbook autoWorkbook;
    /**
     * This field is null for write mode.
     */
    @Nullable
    protected KeelSheetsReaderType sheetsReaderType;

    /**
     * @param workbook The generated POI Workbook Implementation.
     * @since 3.0.20
     */
    protected KeelSheets(@Nullable KeelSheetsReaderType sheetsReaderType, @NotNull Workbook workbook) {
        this(sheetsReaderType, workbook, false);
    }

    /**
     * Create a new Sheets.
     */
    protected KeelSheets() {
        this(null, null, false);
    }

    /**
     * Open an existed workbook or create. Not use stream-write mode by default.
     *
     * @param workbook if null, create a new Sheets; otherwise, use it.
     * @since 3.1.3
     */
    protected KeelSheets(@Nullable KeelSheetsReaderType sheetsReaderType, @Nullable Workbook workbook, boolean withFormulaEvaluator) {
        this.sheetsReaderType = sheetsReaderType;
        autoWorkbook = Objects.requireNonNullElseGet(workbook, XSSFWorkbook::new);
        if (withFormulaEvaluator) {
            formulaEvaluator = autoWorkbook.getCreationHelper().createFormulaEvaluator();
        } else {
            formulaEvaluator = null;
        }
    }

    /**
     * @since 4.0.2
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
                                         } catch (IOException e) {
                                             try {
                                                 workbook = new HSSFWorkbook(inputStream);
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
     * @since 4.0.2
     */
    public static <T> Future<T> useSheets(@NotNull SheetsCreateOptions sheetsCreateOptions,
                                          @NotNull Function<KeelSheets, Future<T>> usage) {
        return Future.succeededFuture()
                     .compose(v -> {
                         KeelSheets keelSheets;
                         if (sheetsCreateOptions.isUseXlsx()) {
                             keelSheets = new KeelSheets(null, new XSSFWorkbook(), sheetsCreateOptions.isWithFormulaEvaluator());
                             if (sheetsCreateOptions.isUseStreamWriting()) {
                                 keelSheets.useStreamWrite();
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
     * @since 4.0.2 became private
     */
    private KeelSheets useStreamWrite() {
        if (autoWorkbook instanceof XSSFWorkbook) {
            autoWorkbook = new SXSSFWorkbook((XSSFWorkbook) autoWorkbook);
        } else {
            throw new IllegalStateException("Now autoWorkbook is not an instance of XSSFWorkbook.");
        }
        return this;
    }

    public KeelSheet generateReaderForSheet(@NotNull String sheetName) {
        return this.generateReaderForSheet(sheetName, true);
    }

    /**
     * @since 3.1.4
     */
    public KeelSheet generateReaderForSheet(@NotNull String sheetName, boolean parseFormulaCellToValue) {
        var sheet = this.getWorkbook().getSheet(sheetName);
        ValueBox<FormulaEvaluator> formulaEvaluatorValueBox = new ValueBox<>();
        if (parseFormulaCellToValue) {
            formulaEvaluatorValueBox.setValue(this.formulaEvaluator);
        }
        return new KeelSheet(sheetsReaderType, sheet, formulaEvaluatorValueBox);
    }

    public KeelSheet generateReaderForSheet(int sheetIndex) {
        return this.generateReaderForSheet(sheetIndex, true);
    }

    /**
     * @since 3.1.4
     */
    public KeelSheet generateReaderForSheet(int sheetIndex, boolean parseFormulaCellToValue) {
        var sheet = this.getWorkbook().getSheetAt(sheetIndex);
        ValueBox<FormulaEvaluator> formulaEvaluatorValueBox = new ValueBox<>();
        if (parseFormulaCellToValue) {
            formulaEvaluatorValueBox.setValue(this.formulaEvaluator);
        }
        return new KeelSheet(sheetsReaderType, sheet, formulaEvaluatorValueBox);
    }

    public KeelSheet generateWriterForSheet(@NotNull String sheetName, Integer pos) {
        Sheet sheet = this.getWorkbook().createSheet(sheetName);
        if (pos != null) {
            this.getWorkbook().setSheetOrder(sheetName, pos);
        }
        return new KeelSheet(null, sheet, new ValueBox<>(this.formulaEvaluator));
    }

    public KeelSheet generateWriterForSheet(@NotNull String sheetName) {
        return generateWriterForSheet(sheetName, null);
    }

    public int getSheetCount() {
        return autoWorkbook.getNumberOfSheets();
    }

    /**
     * @return Raw Apache POI Workbook instance.
     */
    @NotNull
    public Workbook getWorkbook() {
        return autoWorkbook;
    }

    public void save(OutputStream outputStream) {
        try {
            autoWorkbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void save(File file) {
        try {
            save(new FileOutputStream(file));
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

    public void save(String fileName) {
        save(new File(fileName));
    }

    /**
     * @param completable the promise to signal when close has completed
     * @since 4.1.0
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
