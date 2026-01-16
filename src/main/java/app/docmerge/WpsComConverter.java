package app.docmerge;

public class WpsComConverter extends AbstractDocComConverter {
    private static final int FORMAT_DOCX_WPS = 12;
    private static final int FORMAT_DOCX_ALT = 16;

    public WpsComConverter(PowerShellRunner powerShellRunner) {
        super(powerShellRunner);
    }

    @Override
    public String engineName() {
        return "WPS 文字";
    }

    @Override
    protected String progId() {
        return "kwps.application";
    }

    @Override
    protected int[] saveFormatPriority() {
        return new int[]{FORMAT_DOCX_WPS, FORMAT_DOCX_ALT};
    }
}
