package io.github.sinri.keel.integration.poi.csv;

import io.vertx.core.Future;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
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

import static io.github.sinri.keel.facade.KeelInstance.Keel;

/**
 * As of 4.1.1, implements {@link Closeable}, and deprecates all the asynchronous write methods.
 * <p>
 * Call static method {@link KeelCsvWriter#write(OutputStream, String, Charset, Function)} or
 * {@link KeelCsvWriter#write(OutputStream, Function)} is the recommended usage.
 *
 * @since 3.1.1
 */
public class KeelCsvWriter implements Closeable {
    private final OutputStream outputStream;
    private final AtomicBoolean atLineBeginningRef;
    private final String separator;
    private final Charset charset;

    public KeelCsvWriter(@Nonnull OutputStream outputStream) {
        this(outputStream, ",", StandardCharsets.UTF_8);
    }

    /**
     * @since 4.1.1
     */
    public KeelCsvWriter(@Nonnull OutputStream outputStream, @Nonnull String separator, @Nonnull Charset charset) {
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
            @Nonnull OutputStream outputStream,
            @Nonnull String separator,
            @Nonnull Charset charset,
            @Nonnull Function<KeelCsvWriter, Future<Void>> writeCsvFunc
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
            @Nonnull OutputStream outputStream,
            @Nonnull Function<KeelCsvWriter, Future<Void>> writeCsvFunc
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
    public void writeCell(@Nonnull String cellValue) throws IOException {
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

    private void writeToOutputStream(@Nonnull String anything) throws IOException {
        outputStream.write(anything.getBytes(charset));
    }

    /**
     * Write a new csv row to the output stream.
     * <p>
     * If an incomplete row is already written, a LINE ENDING would be added first to enforce a new row output.
     */
    public void blockWriteRow(@Nonnull List<String> list) throws IOException {
        synchronized (atLineBeginningRef) {
            if (!atLineBeginningRef.get()) {
                writeToOutputStream("\n");
                atLineBeginningRef.set(true);
            }
            List<String> components = new ArrayList<>();
            list.forEach(item -> components.add(quote(item)));
            var line = Keel.stringHelper().joinStringArray(components, separator) + "\n";
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
