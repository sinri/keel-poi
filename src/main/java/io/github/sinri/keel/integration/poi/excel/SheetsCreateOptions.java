package io.github.sinri.keel.integration.poi.excel;

/**
 * Excel 工作簿创建选项类，用于配置创建 Excel 工作簿时的参数。
 *
 * @since 5.0.0
 */
public class SheetsCreateOptions {
    private boolean withFormulaEvaluator = false;
    private boolean useXlsx = true;
    private boolean useStreamWriting = true;

    /**
     * 检查是否使用 XLSX 格式。
     *
     * @return 如果使用 XLSX 格式则返回 true，否则返回 false
     */
    public boolean isUseXlsx() {
        return useXlsx;
    }

    /**
     * 设置是否使用 XLSX 格式。
     *
     * @param useXlsx 是否使用 XLSX 格式
     * @return 当前选项实例，支持链式调用
     */
    public SheetsCreateOptions setUseXlsx(boolean useXlsx) {
        this.useXlsx = useXlsx;
        return this;
    }

    /**
     * 检查是否使用流式写入。
     *
     * @return 如果使用流式写入则返回 true，否则返回 false
     */
    public boolean isUseStreamWriting() {
        return useStreamWriting;
    }

    /**
     * 设置是否使用流式写入。
     *
     * @param useStreamWriting 是否使用流式写入
     * @return 当前选项实例，支持链式调用
     */
    public SheetsCreateOptions setUseStreamWriting(boolean useStreamWriting) {
        this.useStreamWriting = useStreamWriting;
        return this;
    }

    /**
     * 检查是否启用公式求值器。
     *
     * @return 如果启用公式求值器则返回 true，否则返回 false
     */
    public boolean isWithFormulaEvaluator() {
        return withFormulaEvaluator;
    }

    /**
     * 设置是否启用公式求值器。
     *
     * @param withFormulaEvaluator 是否启用公式求值器
     * @return 当前选项实例，支持链式调用
     */
    public SheetsCreateOptions setWithFormulaEvaluator(boolean withFormulaEvaluator) {
        this.withFormulaEvaluator = withFormulaEvaluator;
        return this;
    }
}
