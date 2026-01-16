package app.docmerge;

import java.io.IOException;

public class DocComConverterResolver {
    private final DocComConverterSelector selector;

    public DocComConverterResolver(DocComConverterSelector selector) {
        this.selector = selector;
    }

    public Resolution resolve(DocConverterMode mode, boolean forceRefresh) {
        DocComConverterSelector.ProbeSummary summary = selector.probeAll(forceRefresh);
        DocComConverterSelector.EngineStatus word = summary.word();
        DocComConverterSelector.EngineStatus wps = summary.wps();
        DocComConverterSelector.Selection selection = null;
        String errorMessage = null;

        switch (mode) {
            case AUTO -> {
                if (word.available()) {
                    selection = new DocComConverterSelector.Selection(selector.wordConverter(), word);
                } else if (wps.available()) {
                    selection = new DocComConverterSelector.Selection(selector.wpsConverter(), wps);
                } else {
                    errorMessage = "未检测到 Word 或 WPS，当前模式为自动（Word优先）";
                }
            }
            case WORD_ONLY -> {
                if (word.available()) {
                    selection = new DocComConverterSelector.Selection(selector.wordConverter(), word);
                } else {
                    errorMessage = "未检测到 Word，当前模式仅 Word";
                }
            }
            case WPS_ONLY -> {
                if (wps.available()) {
                    selection = new DocComConverterSelector.Selection(selector.wpsConverter(), wps);
                } else {
                    errorMessage = "未检测到 WPS，当前模式仅 WPS";
                }
            }
            default -> errorMessage = "未检测到可用的 DOC 转换引擎";
        }
        return new Resolution(mode, summary, selection, errorMessage);
    }

    public DocComConverterSelector.Selection requireSelection(DocConverterMode mode, boolean forceRefresh) throws IOException {
        Resolution resolution = resolve(mode, forceRefresh);
        if (resolution.selection() == null) {
            throw new IOException(resolution.errorMessage() == null ? "未检测到可用的 DOC 转换引擎" : resolution.errorMessage());
        }
        return resolution.selection();
    }

    public record Resolution(DocConverterMode mode,
                             DocComConverterSelector.ProbeSummary probeSummary,
                             DocComConverterSelector.Selection selection,
                             String errorMessage) {
        public String modeLabel() {
            return mode.getLabel();
        }
    }
}
