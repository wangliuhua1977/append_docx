package com.example.docmerger.service;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.prefs.Preferences;

public final class OrderPersistence {
    private static final String KEY_CUSTOM_ORDER = "customOrder";
    private static final String KEY_SOURCE_DIR = "customOrderSourceDir";

    private final Preferences preferences;

    public OrderPersistence(Class<?> scope) {
        this.preferences = Preferences.userNodeForPackage(scope);
    }

    public void saveCustomOrder(String sourceDir, List<String> stableIds) {
        if (stableIds == null || stableIds.isEmpty()) {
            clearCustomOrder();
            return;
        }
        preferences.put(KEY_SOURCE_DIR, sourceDir);
        preferences.put(KEY_CUSTOM_ORDER, String.join("\n", stableIds));
    }

    public Optional<StoredOrder> loadCustomOrder() {
        String storedDir = preferences.get(KEY_SOURCE_DIR, "");
        String raw = preferences.get(KEY_CUSTOM_ORDER, "");
        if (storedDir.isEmpty() || raw.isEmpty()) {
            return Optional.empty();
        }
        List<String> ids = new ArrayList<>();
        for (String line : raw.split("\R")) {
            String trimmed = line.trim();
            if (!trimmed.isEmpty()) {
                ids.add(trimmed);
            }
        }
        if (ids.isEmpty()) {
            return Optional.empty();
        }
        return Optional.of(new StoredOrder(storedDir, ids));
    }

    public void clearCustomOrder() {
        preferences.remove(KEY_CUSTOM_ORDER);
        preferences.remove(KEY_SOURCE_DIR);
    }

    public record StoredOrder(String sourceDir, List<String> stableIds) {
    }
}
