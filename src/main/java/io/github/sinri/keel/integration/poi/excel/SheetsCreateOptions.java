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

    public boolean isUseXlsx() {
        return useXlsx;
    }

    public SheetsCreateOptions setUseXlsx(boolean useXlsx) {
        this.useXlsx = useXlsx;
        return this;
    }

    public boolean isUseStreamWriting() {
        return useStreamWriting;
    }

    public SheetsCreateOptions setUseStreamWriting(boolean useStreamWriting) {
        this.useStreamWriting = useStreamWriting;
        return this;
    }

    public boolean isWithFormulaEvaluator() {
        return withFormulaEvaluator;
    }

    public SheetsCreateOptions setWithFormulaEvaluator(boolean withFormulaEvaluator) {
        this.withFormulaEvaluator = withFormulaEvaluator;
        return this;
    }
}
