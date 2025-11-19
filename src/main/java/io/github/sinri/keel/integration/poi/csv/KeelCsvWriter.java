package io.github.sinri.keel.integration.poi.csv;

import io.vertx.core.Future;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

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
public class KeelCsvWriter implements Closeable {
    private final OutputStream outputStream;
    private final AtomicBoolean atLineBeginningRef;
    private final String separator;
    private final Charset charset;

    public KeelCsvWriter(@NotNull OutputStream outputStream) {
        this(outputStream, ",", StandardCharsets.UTF_8);
    }

    /**
     * @since 4.1.1
     */
    public KeelCsvWriter(@NotNull OutputStream outputStream, @NotNull String separator, @NotNull Charset charset) {
        this.outputStream = outputStream;
        this.separator = separator;
        this.charset = charset;
        this.atLineBeginningRef = new AtomicBoolean(true);
    }

    /**
     * @param outputStream the output stream
     * @param separator    the separator
     * @param writeCsvFunc the asynchronous function to write csv with a generated {@link KeelCsvWriter} instance, which
     *                     is ensured to be closed automatically in the ending.
     * @since 4.1.1
     */
    public static Future<Void> write(
            @NotNull OutputStream outputStream,
            @NotNull String separator,
            @NotNull Charset charset,
            @NotNull Function<KeelCsvWriter, Future<Void>> writeCsvFunc
    ) {
        AtomicReference<KeelCsvWriter> ref = new AtomicReference<>();
        return Future.succeededFuture()
                     .compose(v -> {
                         var x = new KeelCsvWriter(outputStream, separator, charset);
                         ref.set(x);
                         return Future.succeededFuture();
                     })
                     .compose(v -> {
                         KeelCsvWriter keelCsvWriter = ref.get();
                         Objects.requireNonNull(keelCsvWriter);
                         return writeCsvFunc.apply(ref.get());
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
     * Run {@link KeelCsvWriter#write(OutputStream, String, Charset, Function)} with separator {@code ","}.
     *
     * @since 4.1.1
     */
    public static Future<Void> write(
            @NotNull OutputStream outputStream,
            @NotNull Function<KeelCsvWriter, Future<Void>> writeCsvFunc
    ) {
        return write(outputStream, ",", StandardCharsets.UTF_8, writeCsvFunc);
    }

    /**
     * Write a quoted string and a separator to the output stream as a csv cell.
     * <p> This method would write a comma as suffix, so the number of columns would be one more than the actual size.
     * If you care about this, avoid use this method and use {@link KeelCsvWriter#blockWriteRow(List)} instead.
     *
     * @since 4.1.1
     */
    public void writeCell(@NotNull String cellValue) throws IOException {
        synchronized (atLineBeginningRef) {
            writeToOutputStream(quote(cellValue) + separator);
            atLineBeginningRef.set(false);
        }
    }

    /**
     * Write a new line to the output stream.
     * <p>
     * It is not recommended to use this method along with {@link KeelCsvWriter#blockWriteRow(List)}.
     *
     * @since 4.1.1
     */
    public void writeRowEnding() throws IOException {
        synchronized (atLineBeginningRef) {
            writeToOutputStream("\n");
            atLineBeginningRef.set(true);
        }
    }

    private void writeToOutputStream(@NotNull String anything) throws IOException {
        outputStream.write(anything.getBytes(charset));
    }

    /**
     * Write a new csv row to the output stream.
     * <p>
     * If an incomplete row is already written, a LINE ENDING would be added first to enforce a new row output.
     */
    public void blockWriteRow(@NotNull List<String> list) throws IOException {
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

    @Override
    public void close() throws IOException {
        this.outputStream.close();
    }
}
