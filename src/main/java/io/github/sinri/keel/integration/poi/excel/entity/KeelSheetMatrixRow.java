package io.github.sinri.keel.integration.poi.excel.entity;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.math.BigDecimal;
import java.util.List;

/**
 * Excel 表格矩阵行类，表示 Excel 表格矩阵中的一行数据。
 * 该类设计为可被重写以实现自定义的行读取器，并使用 BigDecimal 处理单元格的数值。
 *
 * @since 5.0.0
 */
public class KeelSheetMatrixRow {
    private final List<String> rawRow;

    public KeelSheetMatrixRow(List<String> rawRow) {
        this.rawRow = rawRow;
    }

    @NotNull
    public String readValue(int i) {
        return rawRow.get(i);
    }

    /**
     * @since 3.1.1
     */
    public BigDecimal readValueToBigDecimal(int i) {
        return new BigDecimal(readValue(i));
    }

    /**
     * @since 3.1.1
     */
    public BigDecimal readValueToBigDecimalStrippedTrailingZeros(int i) {
        return new BigDecimal(readValue(i)).stripTrailingZeros();
    }

    @Nullable
    public Integer readValueToInteger(int i) {
        try {
            return readValueToBigDecimal(i).intValueExact();
        } catch (ArithmeticException arithmeticException) {
            return null;
        }
//        double v = readValueToDouble(i);
//        return (int) v;
    }

    @Nullable
    public Long readValueToLong(int i) {
        try {
            return readValueToBigDecimal(i).longValueExact();
        } catch (ArithmeticException arithmeticException) {
            return null;
        }
//        double v = readValueToDouble(i);
//        return (long) v;
    }

    public double readValueToDouble(int i) {
        return readValueToBigDecimal(i).doubleValue();
//        return Double.parseDouble(readValue(i));
    }
}
