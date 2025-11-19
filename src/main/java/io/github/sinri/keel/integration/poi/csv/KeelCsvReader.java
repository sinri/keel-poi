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
 * CSV 文件读取器，提供对 CSV 文件的解析和读取功能。
 * <p>
 * 推荐使用静态方法 {@link KeelCsvReader#read(InputStream, Charset, String, Function)}。
 * <p>
 *     TODO: 将来实现 {@link Iterator} 接口，并移除所有异步方法。
 *
 * @since 5.0.0
 */
public class KeelCsvReader implements Closeable {
    private final BufferedReader br;
    private final String separator;

    /**
     * 构造函数，使用指定的 BufferedReader 和分隔符创建 CSV 读取器。
     *
     * @param br        用于读取 CSV 数据的 BufferedReader
     * @param separator CSV 文件中使用的分隔符
     */
    public KeelCsvReader(@NotNull BufferedReader br, @NotNull String separator) {
        this.br = br;
        this.separator = separator;
    }

    public KeelCsvReader(@NotNull InputStream inputStream, Charset charset) {
        this(inputStream, charset, ",");
    }

    /**
     * 构造函数，使用指定的输入流、字符集和分隔符创建 CSV 读取器。
     *
     * @param inputStream 用于读取 CSV 数据的输入流
     * @param charset     CSV 文件的字符集
     * @param separator   CSV 文件中使用的分隔符
     */
    public KeelCsvReader(@NotNull InputStream inputStream, @NotNull Charset charset, @NotNull String separator) {
        this(new BufferedReader(new InputStreamReader(inputStream, charset)), separator);
    }

    public KeelCsvReader(@NotNull BufferedReader br) {
        this(br, ",");
    }

    /**
     * 使用指定的输入流、字符集和分隔符读取 CSV 数据，并通过提供的函数处理数据。
     * 该方法会自动管理 CSV 读取器的生命周期，确保在操作完成后关闭读取器。
     *
     * @param inputStream 用于读取 CSV 数据的输入流
     * @param charset     CSV 文件的字符集
     * @param separator   CSV 文件中使用的分隔符
     * @param readFunc    用于读取和处理 CSV 数据的函数
     * @return 表示操作完成的 Future
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
     * 从 CSV 源中读取并解析下一行数据。
     *
     * @return 解析后的 CSV 行对象，如果没有更多行则返回 null
     * @throws IOException 当 CSV 源发生 IO 异常时抛出
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

    /**
     * 关闭 CSV 读取器，释放相关资源。
     *
     * @throws IOException 当关闭过程中发生 IO 异常时抛出
     */
    @Override
    public void close() throws IOException {
        if (br != null) br.close();
    }
}
