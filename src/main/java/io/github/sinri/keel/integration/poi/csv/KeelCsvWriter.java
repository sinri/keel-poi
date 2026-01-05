package io.github.sinri.keel.integration.poi.csv;

import io.vertx.core.Future;
import org.jspecify.annotations.NullMarked;
import org.jspecify.annotations.Nullable;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;


/**
 * CSV 文件写入器，提供对 CSV 文件的生成和写入功能。
 * <p>
 * 推荐使用静态方法 {@link KeelCsvWriter#write(OutputStream, String, Charset, Function)} 或
 * {@link KeelCsvWriter#write(OutputStream, Function)}。
 *
 * @since 5.0.0
 */
@NullMarked
public class KeelCsvWriter implements Closeable {
    private final OutputStream outputStream;
    private final AtomicBoolean atLineBeginningRef;
    private final String separator;
    private final Charset charset;

    public KeelCsvWriter(OutputStream outputStream) {
        this(outputStream, ",", StandardCharsets.UTF_8);
    }

    /**
     * 构造函数，使用指定的输出流、分隔符和字符集创建 CSV 写入器。
     *
     * @param outputStream 用于写入 CSV 数据的输出流
     * @param separator    CSV 文件中使用的分隔符
     * @param charset      CSV 文件的字符集
     */
    public KeelCsvWriter(OutputStream outputStream, String separator, Charset charset) {
        this.outputStream = outputStream;
        this.separator = separator;
        this.charset = charset;
        this.atLineBeginningRef = new AtomicBoolean(true);
    }

    /**
     * 使用指定的输出流、分隔符和字符集写入 CSV 数据，并通过提供的函数处理写入操作。
     * 该方法会自动管理 CSV 写入器的生命周期，确保在操作完成后关闭写入器。
     *
     * @param outputStream 用于写入 CSV 数据的输出流
     * @param separator    CSV 文件中使用的分隔符
     * @param charset      CSV 文件的字符集
     * @param writeCsvFunc 用于写入 CSV 数据的函数
     * @return 表示操作完成的 Future
     */
    public static Future<Void> write(
            OutputStream outputStream,
            String separator,
            Charset charset,
            Function<KeelCsvWriter, Future<Void>> writeCsvFunc
    ) {
        AtomicReference<@Nullable KeelCsvWriter> ref = new AtomicReference<>();
        return Future.succeededFuture()
                     .compose(v -> {
                         var x = new KeelCsvWriter(outputStream, separator, charset);
                         ref.set(x);
                         return Future.succeededFuture();
                     })
                     .compose(v -> {
                         KeelCsvWriter keelCsvWriter = ref.get();
                         return writeCsvFunc.apply(Objects.requireNonNull(keelCsvWriter));
                     })
                     .eventually(() -> {
                         KeelCsvWriter keelCsvWriter = ref.get();
                         if (keelCsvWriter != null) {
                             try {
                                 keelCsvWriter.close();
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
     * 使用默认分隔符（逗号）和字符集写入 CSV 数据，并通过提供的函数处理写入操作。
     * 该方法会自动管理 CSV 写入器的生命周期，确保在操作完成后关闭写入器。
     *
     * @param outputStream 用于写入 CSV 数据的输出流
     * @param writeCsvFunc 用于写入 CSV 数据的函数
     * @return 表示操作完成的 Future
     */
    public static Future<Void> write(
            OutputStream outputStream,
            Function<KeelCsvWriter, Future<Void>> writeCsvFunc
    ) {
        return write(outputStream, ",", StandardCharsets.UTF_8, writeCsvFunc);
    }

    /**
     * 将带引号的字符串和分隔符写入输出流作为 CSV 单元格。
     * <p> 该方法会在单元格后写入分隔符，因此列数会比实际大小多一列。
     * 如果您关心这一点，请避免使用此方法，而使用 {@link KeelCsvWriter#blockWriteRow(List)}。
     *
     * @param cellValue 要写入的单元格值
     * @throws IOException 当写入过程中发生 IO 异常时抛出
     */
    public void writeCell(String cellValue) throws IOException {
        synchronized (atLineBeginningRef) {
            writeToOutputStream(quote(cellValue) + separator);
            atLineBeginningRef.set(false);
        }
    }

    /**
     * 将新行写入输出流。
     * <p>
     * 不建议将此方法与 {@link KeelCsvWriter#blockWriteRow(List)} 一起使用。
     *
     * @throws IOException 当写入过程中发生 IO 异常时抛出
     */
    public void writeRowEnding() throws IOException {
        synchronized (atLineBeginningRef) {
            writeToOutputStream("\n");
            atLineBeginningRef.set(true);
        }
    }

    private void writeToOutputStream(String anything) throws IOException {
        outputStream.write(anything.getBytes(charset));
    }

    /**
     * 将新的 CSV 行写入输出流。
     * <p>
     * 如果已经写入了不完整的行，则会首先添加行结束符以确保输出新行。
     *
     * @param list 要写入的行数据列表
     * @throws IOException 当写入过程中发生 IO 异常时抛出
     */
    public void blockWriteRow(List<String> list) throws IOException {
        synchronized (atLineBeginningRef) {
            if (!atLineBeginningRef.get()) {
                writeToOutputStream("\n");
                atLineBeginningRef.set(true);
            }
            List<String> components = new ArrayList<>();
            list.forEach(item -> components.add(quote(item)));
            var line = String.join(separator, components) + "\n";
            writeToOutputStream(line);
        }
    }

    private String quote(@Nullable String s) {
        if (s == null) {
            s = "";
        }
        if (s.contains("\"") || s.contains("\n") || s.contains(separator)) {
            return "\"" + s.replaceAll("\"", "\"\"") + "\"";
        } else {
            return s;
        }
    }

    /**
     * 关闭 CSV 写入器，释放相关资源。
     *
     * @throws IOException 当关闭过程中发生 IO 异常时抛出
     */
    @Override
    public void close() throws IOException {
        this.outputStream.close();
    }
}
