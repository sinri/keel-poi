package io.github.sinri.keel.integration.poi.excel;

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
