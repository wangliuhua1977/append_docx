package app.docmerge;

public enum DocConverterMode {
    AUTO("自动（Word优先）"),
    WORD_ONLY("仅 Word"),
    WPS_ONLY("仅 WPS");

    private final String label;

    DocConverterMode(String label) {
        this.label = label;
    }

    public String getLabel() {
        return label;
    }

    @Override
    public String toString() {
        return label;
    }

    public static DocConverterMode fromConfig(String value) {
        if (value == null || value.isBlank()) {
            return AUTO;
        }
        try {
            return DocConverterMode.valueOf(value.trim().toUpperCase());
        } catch (IllegalArgumentException ex) {
            return AUTO;
        }
    }
}
