package app.docmerge;

public class DocComConverterSelector {
    private final MsWordComConverter wordConverter;
    private final WpsComConverter wpsConverter;
    private ProbeSummary cachedSummary;
    private long lastProbeMillis;
    private static final long CACHE_TTL_MILLIS = 60_000L;

    public DocComConverterSelector() {
        PowerShellRunner runner = new PowerShellRunner();
        this.wordConverter = new MsWordComConverter(runner);
        this.wpsConverter = new WpsComConverter(runner);
    }

    public ProbeSummary probeAll() {
        return probeAll(false);
    }

    public synchronized ProbeSummary probeAll(boolean forceRefresh) {
        if (!forceRefresh && cachedSummary != null && !isExpired()) {
            return cachedSummary;
        }
        EngineStatus word = probeWord();
        EngineStatus wps = probeWps();
        Selection selection = selectAvailable(word, wps);
        cachedSummary = new ProbeSummary(word, wps, selection);
        lastProbeMillis = System.currentTimeMillis();
        return cachedSummary;
    }

    public EngineStatus probeWord() {
        DocComProbeResult result = wordConverter.probe();
        return new EngineStatus("Word", wordConverter.engineName(), result.available(), result.message(), result.stdout(), result.stderr(), result.exitCode());
    }

    public EngineStatus probeWps() {
        DocComProbeResult result = wpsConverter.probe();
        return new EngineStatus("WPS", wpsConverter.engineName(), result.available(), result.message(), result.stdout(), result.stderr(), result.exitCode());
    }

    public MsWordComConverter wordConverter() {
        return wordConverter;
    }

    public WpsComConverter wpsConverter() {
        return wpsConverter;
    }

    private Selection selectAvailable(EngineStatus word, EngineStatus wps) {
        if (word.available()) {
            return new Selection(wordConverter, word);
        }
        if (wps.available()) {
            return new Selection(wpsConverter, wps);
        }
        return null;
    }

    public record EngineStatus(String label,
                               String engineName,
                               boolean available,
                               String message,
                               String stdout,
                               String stderr,
                               int exitCode) {
    }

    public record Selection(DocComConverter converter, EngineStatus status) {
    }

    public record ProbeSummary(EngineStatus word,
                               EngineStatus wps,
                               Selection selection) {
        public boolean anyAvailable() {
            return (word != null && word.available()) || (wps != null && wps.available());
        }

        public String currentEngineLabel() {
            if (selection == null) {
                return "无";
            }
            return selection.status().engineName();
        }
    }

    private boolean isExpired() {
        if (lastProbeMillis <= 0) {
            return true;
        }
        return System.currentTimeMillis() - lastProbeMillis > CACHE_TTL_MILLIS;
    }
}
