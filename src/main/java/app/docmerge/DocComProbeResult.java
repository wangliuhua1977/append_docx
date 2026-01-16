package app.docmerge;

public record DocComProbeResult(String engineName,
                                boolean available,
                                String message,
                                String stdout,
                                String stderr,
                                int exitCode) {
    public static DocComProbeResult unavailable(String engineName, String message) {
        return new DocComProbeResult(engineName, false, message, "", message, -1);
    }
}
