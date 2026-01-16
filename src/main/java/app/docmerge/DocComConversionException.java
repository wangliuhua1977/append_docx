package app.docmerge;

import java.io.IOException;

public class DocComConversionException extends IOException {
    private final String failedInput;
    private final String stdout;
    private final String stderr;
    private final int exitCode;

    public DocComConversionException(String message, String failedInput, String stdout, String stderr, int exitCode) {
        super(message);
        this.failedInput = failedInput;
        this.stdout = stdout;
        this.stderr = stderr;
        this.exitCode = exitCode;
    }

    public String getFailedInput() {
        return failedInput;
    }

    public String getStdout() {
        return stdout;
    }

    public String getStderr() {
        return stderr;
    }

    public int getExitCode() {
        return exitCode;
    }
}
