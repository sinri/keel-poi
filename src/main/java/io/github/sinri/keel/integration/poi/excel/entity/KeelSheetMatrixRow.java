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

    /**
     * 构造函数，使用指定的原始行数据创建表格矩阵行实例。
     *
     * @param rawRow 原始行数据列表
     * @since 5.0.0
     */
    public KeelSheetMatrixRow(List<String> rawRow) {
        this.rawRow = rawRow;
    }

    /**
     * 读取指定索引处的单元格值。
     *
     * @param i 单元格索引
     * @return 指定索引处的单元格值
     * @since 5.0.0
     */
    @NotNull
    public String readValue(int i) {
        return rawRow.get(i);
    }

    /**
     * 读取指定索引处的单元格值并转换为 BigDecimal。
     *
     * @param i 单元格索引
     * @return 转换后的 BigDecimal 值
     * @since 5.0.0
     */
    public BigDecimal readValueToBigDecimal(int i) {
        return new BigDecimal(readValue(i));
    }

    /**
     * 读取指定索引处的单元格值并转换为 BigDecimal，同时去除尾随零。
     *
     * @param i 单元格索引
     * @return 转换后的 BigDecimal 值，已去除尾随零
     * @since 5.0.0
     */
    public BigDecimal readValueToBigDecimalStrippedTrailingZeros(int i) {
        return new BigDecimal(readValue(i)).stripTrailingZeros();
    }

    /**
     * 读取指定索引处的单元格值并转换为整数。
     * 如果转换失败或值超出整数范围，则返回 null。
     *
     * @param i 单元格索引
     * @return 转换后的整数值，如果转换失败则返回 null
     * @since 5.0.0
     */
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

    /**
     * 读取指定索引处的单元格值并转换为长整数。
     * 如果转换失败或值超出长整数范围，则返回 null。
     *
     * @param i 单元格索引
     * @return 转换后的长整数值，如果转换失败则返回 null
     * @since 5.0.0
     */
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

    /**
     * 读取指定索引处的单元格值并转换为双精度浮点数。
     *
     * @param i 单元格索引
     * @return 转换后的双精度浮点数值
     * @since 5.0.0
     */
    public double readValueToDouble(int i) {
        return readValueToBigDecimal(i).doubleValue();
//        return Double.parseDouble(readValue(i));
    }
}
