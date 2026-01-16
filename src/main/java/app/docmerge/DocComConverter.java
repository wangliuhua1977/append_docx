package app.docmerge;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

public interface DocComConverter {
    String engineName();

    boolean isAvailable();

    List<Path> convertBatch(List<Path> docFiles, Path tempDir) throws IOException, DocComConversionException;
}
