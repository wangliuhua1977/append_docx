package app.docmerge;

public class DocComConverterSelector {
    private final MsWordComConverter wordConverter;
    private final WpsComConverter wpsConverter;

    public DocComConverterSelector() {
        PowerShellRunner runner = new PowerShellRunner();
        this.wordConverter = new MsWordComConverter(runner);
        this.wpsConverter = new WpsComConverter(runner);
    }

    public ProbeSummary probeAll() {
        EngineStatus word = probeWord();
        EngineStatus wps = probeWps();
        Selection selection = selectAvailable(word, wps);
        return new ProbeSummary(word, wps, selection);
    }

    public EngineStatus probeWord() {
        DocComProbeResult result = wordConverter.probe();
        return new EngineStatus("Word", wordConverter.engineName(), result.available(), result.message(), result.stdout(), result.stderr(), result.exitCode());
    }

    public EngineStatus probeWps() {
        DocComProbeResult result = wpsConverter.probe();
        return new EngineStatus("WPS", wpsConverter.engineName(), result.available(), result.message(), result.stdout(), result.stderr(), result.exitCode());
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
}
