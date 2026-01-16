package app.docmerge;

public class MsWordComConverter extends AbstractDocComConverter {
    private static final int FORMAT_DOCX_WORD = 16;
    private static final int FORMAT_DOCX_ALT = 12;

    public MsWordComConverter(PowerShellRunner powerShellRunner) {
        super(powerShellRunner);
    }

    @Override
    public String engineName() {
        return "Microsoft Word";
    }

    @Override
    protected String progId() {
        return "Word.Application";
    }

    @Override
    protected int[] saveFormatPriority() {
        return new int[]{FORMAT_DOCX_WORD, FORMAT_DOCX_ALT};
    }
}
