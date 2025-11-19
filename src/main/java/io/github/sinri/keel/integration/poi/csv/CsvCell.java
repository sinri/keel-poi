package io.github.sinri.keel.integration.poi.csv;


import io.github.sinri.keel.core.utils.value.ValueBox;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.math.BigDecimal;

/**
 * As of 4.1.1, number parsing is executed lazily when requested, not
 * precomputed.
 *
 * @since 3.1.1
 */
public class CsvCell {
    private final @Nullable String string;
    private final ValueBox<BigDecimal> bigDecimalValueBox = new ValueBox<>();

    /**
     * As of 4.1.1, do not attempt to parse numbers in the constructor.
     * <p>
     * 
     * @param s the CSV cell value
     */
    public CsvCell(@Nullable String s) {
        this.string = s;
    }

    /**
     * This method follows these steps:
     * <p>
     * First, if this cell is null, return false.
     * <p>
     * Second, if no cached result is present, try to parse the non-null value as a
     * number and store it in {@link #bigDecimalValueBox}. Skip this step if already
     * cached.
     * <p>
     * Finally, return true if the value has been parsed to a number, otherwise
     * false.
     * <p>
     * 
     * @return whether this cell value can be parsed to a number
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
     * <p>
     * 
     * @return whether this cell is non-null and empty
     */
    public boolean isEmpty() {
        return this.string != null && this.string.isEmpty();
    }

    /**
     * <p>
     * 
     * @return whether this cell is null
     */
    public boolean isNull() {
        return this.string == null;
    }

    /**
     * As of 4.1.1, number parsing is executed lazily when called, not precomputed.
     * <p>
     * 
     * @return the parsed number of this cell value, or null when this cell value is
     *         null
     * @throws NumberFormatException when the cell value cannot be parsed to a
     *                               number
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
     * Get the numeric value of this cell, or the default value if the cell value is
     * not a number.
     * 
     * @param defaultValue the default value to return if the cell value is not a
     *                     number
     * @return the parsed number of this cell value, or the default value if the
     *         cell value is not a number
     * @since 4.1.1
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
     * Get the numeric value of this cell, or the default value if the cell value is
     * not a number.
     * 
     * @param defaultValue the default value to return if the cell value is not a
     *                     number
     * @return the parsed number of this cell value, or the default value if the
     *         cell value is not a number
     * @since 4.1.1
     */
    public BigDecimal getNumberOrElse(int defaultValue) {
        return getNumberOrElse(BigDecimal.valueOf(defaultValue));
    }
    /**
     * Get the numeric value of this cell, or the default value if the cell value is
     * not a number.
     * 
     * @param defaultValue the default value to return if the cell value is not a
     *                     number
     * @return the parsed number of this cell value, or the default value if the
     *         cell value is not a number
     * @since 4.1.1
     */
    public BigDecimal getNumberOrElse(long defaultValue) {
        return getNumberOrElse(BigDecimal.valueOf(defaultValue));
    }
    /**
     * Get the numeric value of this cell, or the default value if the cell value is
     * not a number.
     * 
     * @param defaultValue the default value to return if the cell value is not a
     *                     number
     * @return the parsed number of this cell value, or the default value if the
     *         cell value is not a number
     * @since 4.1.1
     */
    public BigDecimal getNumberOrElse(double defaultValue) {
        return getNumberOrElse(BigDecimal.valueOf(defaultValue));
    }
    /**
     * Get the numeric value of this cell, or the default value if the cell value is
     * not a number.
     * 
     * @param defaultValue the default value to return if the cell value is not a
     *                     number
     * @return the parsed number of this cell value, or the default value if the
     *         cell value is not a number
     * @since 4.1.1
     */
    public BigDecimal getNumberOrElse(float defaultValue) {
        return getNumberOrElse(BigDecimal.valueOf(defaultValue));
    }

    /**
     * Get the string value of this cell, or null if this cell is null.
     * 
     * @return the string value of this cell, or null if this cell is null
     */
    @Nullable
    public String getString() {
        return string;
    }
}
