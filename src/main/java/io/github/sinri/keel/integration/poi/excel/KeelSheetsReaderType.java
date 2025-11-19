package io.github.sinri.keel.integration.poi.excel;

/**
 * Excel 工作簿读取器类型枚举，定义了不同的 Excel 文件读取方式。
 *
 * @since 5.0.0
 */
public enum KeelSheetsReaderType {
    XLSX_STREAMING,  // XLSX 流式读取
    XLSX,            // XLSX 标准读取
    XLS              // XLS 标准读取
}
