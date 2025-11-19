package io.github.sinri.keel.integration.poi.excel;

import com.github.pjfanning.xlsx.StreamingReader;
import io.vertx.core.Handler;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.io.File;
import java.io.InputStream;
import java.util.Objects;

/**
 * Excel 工作簿打开选项类，用于配置打开 Excel 工作簿时的参数。
 *
 * @since 5.0.0
 */
public class SheetsOpenOptions {
    private boolean withFormulaEvaluator = false;
    @Nullable
    private File file = null;
    @Nullable
    private StreamingReader.Builder hugeXlsxStreamingReaderBuilder = null;
    @Nullable
    private InputStream inputStream = null;
    @Nullable
    private Boolean useXlsx = null;

    /**
     * 配置读取超大 Excel 文件的系统参数。
     * <p>
     * excel-streaming-reader 在底层使用一些 Apache POI 代码。该代码在处理 xlsx 时使用内存和（或）临时文件来存储临时数据。
     * 对于非常大的文件，您可能希望优先使用临时文件。
     * <p>
     * 使用 StreamingReader.builder() 时，请勿设置 setAvoidTempFiles(true)。您还应考虑调整 POI 设置。
     */
    public static void declareReadingVeryLargeExcelFiles() {
        org.apache.poi.openxml4j.util.ZipInputStreamZipEntrySource.setThresholdBytesForTempFiles(16384); //16KB
        org.apache.poi.openxml4j.opc.ZipPackage.setUseTempFilePackageParts(true);
    }

    boolean isWithFormulaEvaluator() {
        return this.withFormulaEvaluator;
    }

    /**
     * 设置是否启用公式求值器。
     *
     * @param withFormulaEvaluator 是否启用公式求值器
     * @return 当前选项实例，支持链式调用
     */
    public SheetsOpenOptions setWithFormulaEvaluator(boolean withFormulaEvaluator) {
        this.withFormulaEvaluator = withFormulaEvaluator;
        return this;
    }

    /**
     * 获取要打开的 Excel 文件。
     *
     * @return Excel 文件对象，可能为 null
     */
    @Nullable
    public File getFile() {
        return this.file;
    }

    /**
     * 设置要打开的 Excel 文件。
     *
     * @param file 要打开的 Excel 文件
     * @return 当前选项实例，支持链式调用
     */
    public SheetsOpenOptions setFile(@NotNull File file) {
        this.file = file;
        return this;
    }

    /**
     * 通过文件路径设置要打开的 Excel 文件。
     *
     * @param filePath 要打开的 Excel 文件路径
     * @return 当前选项实例，支持链式调用
     */
    public SheetsOpenOptions setFile(@NotNull String filePath) {
        this.file = new File(filePath);
        return this;
    }

    /**
     * 获取要读取的输入流。
     *
     * @return 输入流对象，可能为 null
     */
    @Nullable
    public InputStream getInputStream() {
        return inputStream;
    }

    /**
     * 设置要读取的输入流。
     *
     * @param inputStream 要读取的输入流
     * @return 当前选项实例，支持链式调用
     */
    public SheetsOpenOptions setInputStream(@NotNull InputStream inputStream) {
        this.inputStream = inputStream;
        return this;
    }

    /**
     * 检查是否使用超大 XLSX 流式读取。
     *
     * @return 如果使用超大 XLSX 流式读取则返回 true，否则返回 false
     */
    public boolean isUseHugeXlsxStreamReading() {
        return this.hugeXlsxStreamingReaderBuilder != null;
    }

    /**
     * 获取超大 XLSX 流式读取构建器。
     * <p>
     * 在使用此方法获取构建器之前，请先通过 {@link SheetsOpenOptions#isUseHugeXlsxStreamReading()} 方法检查是否启用了超大 XLSX 流式读取。
     *
     * @return 超大 XLSX 流式读取构建器
     */
    @NotNull
    public StreamingReader.Builder getHugeXlsxStreamingReaderBuilder() {
        Objects.requireNonNull(this.hugeXlsxStreamingReaderBuilder);
        return hugeXlsxStreamingReaderBuilder;
    }

    /**
     * 设置超大 XLSX 文件的流式读取构建器。
     * <p>
     * 此方法用于使用 `pjfanning::excel-streaming-reader` 读取超大 XLSX 文件。
     * 您可以在行内随机访问单元格，因为整行会被缓存。但是，无法随机访问行。
     * 由于这是流式实现，在任何给定时间只会在内存中保留少量行。
     * </p>
     * <p>
     * 请考虑处理临时文件共享字符串和临时文件注释。
     * </p>
     *
     * @param streamingReaderBuilderHandler 流式读取构建器处理器
     * @return 当前选项实例，支持链式调用
     * @see <a href="https://github.com/pjfanning/excel-streaming-reader">PJFANNING::ExcelStreamingReader</a>
     */
    public SheetsOpenOptions setHugeXlsxStreamingReaderBuilder(@NotNull Handler<StreamingReader.Builder> streamingReaderBuilderHandler) {
        var hugeXlsxStreamingReaderBuilder = new StreamingReader.Builder();

        // number of rows to keep in memory (defaults to 10)
        hugeXlsxStreamingReaderBuilder.rowCacheSize(32);
        // buffer size (in bytes) to use when reading InputStream to file (defaults to 1024)
        hugeXlsxStreamingReaderBuilder.bufferSize(10240);

        streamingReaderBuilderHandler.handle(hugeXlsxStreamingReaderBuilder);

        this.hugeXlsxStreamingReaderBuilder = hugeXlsxStreamingReaderBuilder;
        this.setUseXlsx(true);
        return this;
    }

    /**
     * 检查是否使用 XLSX 格式。
     *
     * @return XLSX 格式的使用状态，或 null 表示未知
     */
    @Nullable
    public Boolean isUseXlsx() {
        return useXlsx;
    }

    /**
     * 设置是否使用 XLSX 格式。
     *
     * @param useXlsx 要设置的 XLSX 格式使用状态
     * @return 当前选项实例，支持链式调用
     */
    public SheetsOpenOptions setUseXlsx(@NotNull Boolean useXlsx) {
        this.useXlsx = useXlsx;
        return this;
    }
}
