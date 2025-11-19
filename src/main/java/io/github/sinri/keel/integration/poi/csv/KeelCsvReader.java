package io.github.sinri.keel.integration.poi.csv;

import io.vertx.core.Future;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.io.*;
import java.nio.charset.Charset;
import java.util.Iterator;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;

/**
 * As of 4.1.1, implements {@link Closeable}, and deprecates all the asynchronous methods.
 * <p>
 * Call static method {@link KeelCsvReader#read(InputStream, Charset, String, Function)} is the recommended usage.
 * <p>
 *     TODO: implement {@link Iterator} in the future, and remove all the asynchronous methods.
 *
 * @since 3.1.1
 */
public class KeelCsvReader implements Closeable {
    private final BufferedReader br;
    private final String separator;

    /**
     * @since 4.1.1
     */
    public KeelCsvReader(@NotNull BufferedReader br, @NotNull String separator) {
        this.br = br;
        this.separator = separator;
    }

    public KeelCsvReader(@NotNull InputStream inputStream, Charset charset) {
        this(inputStream, charset, ",");
    }

    /**
     * @since 4.1.1
     */
    public KeelCsvReader(@NotNull InputStream inputStream, @NotNull Charset charset, @NotNull String separator) {
        this(new BufferedReader(new InputStreamReader(inputStream, charset)), separator);
    }

    public KeelCsvReader(@NotNull BufferedReader br) {
        this(br, ",");
    }

    /**
     * @param inputStream the input stream for csv
     * @param charset     the charset used by the csv
     * @param separator   the separator used by the csv
     * @param readFunc    a function to read the csv and process the data with a generated {@link KeelCsvReader}
     *                    instance.
     * @since 4.1.1
     */
    public static Future<Void> read(
            @NotNull InputStream inputStream, @NotNull Charset charset, @NotNull String separator,
            @NotNull Function<KeelCsvReader, Future<Void>> readFunc
    ) {
        AtomicReference<KeelCsvReader> ref = new AtomicReference<>();
        return Future.succeededFuture()
                     .compose(v -> {
                         KeelCsvReader keelCsvReader = new KeelCsvReader(inputStream, charset, separator);
                         ref.set(keelCsvReader);
                         return Future.succeededFuture();
                     })
                     .compose(v -> {
                         KeelCsvReader keelCsvReader = ref.get();
                         Objects.requireNonNull(keelCsvReader);
                         return readFunc.apply(keelCsvReader);
                     })
                     .eventually(() -> {
                         KeelCsvReader keelCsvReader = ref.get();
                         if (keelCsvReader != null) {
                             try {
                                 keelCsvReader.close();
                                 return Future.succeededFuture();
                             } catch (IOException e) {
                                 return Future.failedFuture(e);
                             }
                         } else {
                             return Future.succeededFuture();
                         }
                     });
    }


    /**
     * @return the next row parsed from csv source, or null if no any more rows there.
     * @throws IOException if any IO exceptions occur about the csv source
     */
    @Nullable
    public CsvRow next() throws IOException {
        String line = br.readLine();
        if (line == null) return null;
        return consumeOneLine(null, null, 0, line);
    }

    /**
     * @param row        the uncompleted row instance, if existed
     * @param buffer     the cell content buffer, if existed
     * @param quoterFlag Three options: 0,1,2
     *                   <p> + aaa,bbb
     *                   <p> - 0000000
     *                   <p> + ,"aa""bb",
     *                   <p> - 0111211122
     * @param line       the raw text of the line
     * @return the parsed row instance; may be incompleted during recursion.
     */
    private CsvRow consumeOneLine(@Nullable CsvRow row, @Nullable StringBuilder buffer, int quoterFlag, @NotNull String line) throws IOException {
        if (row == null) {
            row = new CsvRow();
        }
        if (buffer == null) {
            buffer = new StringBuilder();
        } else {
            buffer.append("\n");
        }

        for (int i = 0; i < line.length(); i++) {
            var singleString = line.substring(i, i + 1);
            if (singleString.equals("\"")) {
                if (quoterFlag == 0) {
                    quoterFlag = 1;
                } else if (quoterFlag == 1) {
                    quoterFlag = 2;
                } else {
                    buffer.append(singleString);
                    quoterFlag = 1;
                }
            } else if (singleString.equals(separator)) {
                if (quoterFlag == 0 || quoterFlag == 2) {
                    // buffer to cell
                    row.addCell(new CsvCell(buffer.toString()));
                    quoterFlag = 0;
                    buffer = new StringBuilder();
                } else {
                    buffer.append(singleString);
                }
            } else {
                buffer.append(singleString);
            }
        }

        // now this line ends
        if (quoterFlag == 0 || quoterFlag == 2) {
            // now the row ends within this line
            row.addCell(new CsvCell(buffer.toString()));
            return row;
        } else {
            // this row seems to expend to the next line
            String nextLine = br.readLine();
            if (nextLine == null) {
                // strange: file ending without escape quote, for safety, escape it:
                // let us handle it as quoterFlag is 2
                row.addCell(new CsvCell(buffer.toString()));
                return row;
            }
            return consumeOneLine(row, buffer, quoterFlag, nextLine);
        }
    }

    @Override
    public void close() throws IOException {
        if (br != null) br.close();
    }
}
