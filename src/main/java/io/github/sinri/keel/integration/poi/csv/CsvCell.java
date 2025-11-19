package io.github.sinri.keel.integration.poi.csv;


import io.github.sinri.keel.core.utils.value.ValueBox;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.math.BigDecimal;

/**
 * CSV 单元格类，表示 CSV 文件中的一个单元格。
 * <p>
 * 数字解析在请求时才执行，而不是预先计算。
 *
 * @since 5.0.0
 */
public class CsvCell {
    private final @Nullable String string;
    private final ValueBox<BigDecimal> bigDecimalValueBox = new ValueBox<>();

    /**
     * 构造函数，使用指定的字符串值创建 CSV 单元格。
     * <p>
     * 在构造函数中不尝试解析数字。
     *
     * @param s CSV 单元格的字符串值
     */
    public CsvCell(@Nullable String s) {
        this.string = s;
    }

    /**
     * 判断此单元格的值是否可以解析为数字。
     * <p>
     * 该方法的执行步骤如下：
     * <br>
     * 首先，如果此单元格为 null，则返回 false。
     * <br>
     * 其次，如果没有缓存结果，则尝试将非空值解析为数字并存储在 {@link #bigDecimalValueBox} 中。
     * 如果已缓存，则跳过此步骤。
     * <br>
     * 最后，如果值已解析为数字，则返回 true，否则返回 false。
     *
     * @return 此单元格的值是否可以解析为数字
     */
    public boolean isNumber() {
        if (this.string == null) {
            return false;
        }
        if (!this.bigDecimalValueBox.isValueAlreadySet()) {
            // Not set yet
            try {
                BigDecimal bigDecimal = new BigDecimal(this.string);
                bigDecimalValueBox.setValue(bigDecimal);
                return true;
            } catch (NumberFormatException numberFormatException) {
                bigDecimalValueBox.setValue(null);
                return false;
            }
        } else {
            // Already Set
            return this.bigDecimalValueBox.isValueSetAndNotNull();
        }
    }

    /**
     * 判断此单元格是否非空且为空字符串。
     *
     * @return 此单元格是否非空且为空字符串
     */
    public boolean isEmpty() {
        return this.string != null && this.string.isEmpty();
    }

    /**
     * 判断此单元格是否为 null。
     *
     * @return 此单元格是否为 null
     */
    public boolean isNull() {
        return this.string == null;
    }

    /**
     * 获取此单元格值解析后的数字。
     * <p>
     * 数字解析在调用时才执行，而不是预先计算。
     *
     * @return 此单元格值解析后的数字，如果此单元格值为 null 则返回 null
     * @throws NumberFormatException 当单元格值无法解析为数字时抛出
     */
    @Nullable
    public BigDecimal getNumber() throws NumberFormatException {
        if (this.string == null)
            return null;

        if (!isNumber()) {
            throw new NumberFormatException("Cell value is not a number: " + this.string);
        }

        return bigDecimalValueBox.getValue();
    }

    /**
     * 获取此单元格的数值，如果单元格值不是数字则返回默认值。
     *
     * @param defaultValue 如果单元格值不是数字则返回的默认值
     * @return 此单元格值解析后的数字，或如果单元格值不是数字则返回默认值
     */
    @NotNull
    public BigDecimal getNumberOrElse(@NotNull BigDecimal defaultValue) {
        try {

            var x = getNumber();
            if (x == null)
                return defaultValue;
            return x;
        } catch (NumberFormatException e) {
            return defaultValue;
        }
    }

    /**
     * 获取此单元格的数值，如果单元格值不是数字则返回默认值。
     *
     * @param defaultValue 如果单元格值不是数字则返回的默认值
     * @return 此单元格值解析后的数字，或如果单元格值不是数字则返回默认值
     */
    public BigDecimal getNumberOrElse(int defaultValue) {
        return getNumberOrElse(BigDecimal.valueOf(defaultValue));
    }

    /**
     * 获取此单元格的数值，如果单元格值不是数字则返回默认值。
     *
     * @param defaultValue 如果单元格值不是数字则返回的默认值
     * @return 此单元格值解析后的数字，或如果单元格值不是数字则返回默认值
     */
    public BigDecimal getNumberOrElse(long defaultValue) {
        return getNumberOrElse(BigDecimal.valueOf(defaultValue));
    }

    /**
     * 获取此单元格的数值，如果单元格值不是数字则返回默认值。
     *
     * @param defaultValue 如果单元格值不是数字则返回的默认值
     * @return 此单元格值解析后的数字，或如果单元格值不是数字则返回默认值
     */
    public BigDecimal getNumberOrElse(double defaultValue) {
        return getNumberOrElse(BigDecimal.valueOf(defaultValue));
    }

    /**
     * 获取此单元格的数值，如果单元格值不是数字则返回默认值。
     *
     * @param defaultValue 如果单元格值不是数字则返回的默认值
     * @return 此单元格值解析后的数字，或如果单元格值不是数字则返回默认值
     */
    public BigDecimal getNumberOrElse(float defaultValue) {
        return getNumberOrElse(BigDecimal.valueOf(defaultValue));
    }

    /**
     * 获取此单元格的字符串值，如果此单元格为 null 则返回 null。
     *
     * @return 此单元格的字符串值，或如果此单元格为 null 则返回 null
     */
    @Nullable
    public String getString() {
        return string;
    }
}
