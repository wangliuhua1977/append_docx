package app.docmerge;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

public interface DocComConverter {
    String engineName();

    boolean isAvailable();

    List<Path> convertBatch(List<Path> docFiles, Path tempDir) throws IOException, DocComConversionException;

    default boolean supportsPdfConversion() {
        return false;
    }

    default Path convertPdfToDocx(Path pdfFile, Path tempDir) throws IOException, DocComConversionException {
        throw new DocComConversionException("当前引擎不支持 PDF 转 DOCX",
                pdfFile == null ? "" : pdfFile.toString(),
                "",
                "当前引擎不支持 PDF 转 DOCX",
                -1);
    }
}
