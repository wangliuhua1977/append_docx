package com.example.docmerger.model;

import javax.swing.table.AbstractTableModel;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public final class FileTableModel extends AbstractTableModel {
    private static final String[] COLUMNS = {"File", "Size", "Last Modified"};
    private static final DateTimeFormatter FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")
            .withZone(ZoneId.systemDefault());

    private final List<FileItem> items = new ArrayList<>();

    public void setItems(List<FileItem> newItems) {
        items.clear();
        items.addAll(newItems);
        fireTableDataChanged();
    }

    public List<FileItem> getItems() {
        return Collections.unmodifiableList(items);
    }

    public void moveRows(int[] rows, int targetIndex) {
        if (rows == null || rows.length == 0) {
            return;
        }
        List<FileItem> moving = new ArrayList<>();
        for (int row : rows) {
            moving.add(items.get(row));
        }
        for (int i = rows.length - 1; i >= 0; i--) {
            items.remove(rows[i]);
        }
        int adjustedIndex = targetIndex;
        for (int row : rows) {
            if (row < targetIndex) {
                adjustedIndex--;
            }
        }
        if (adjustedIndex < 0) {
            adjustedIndex = 0;
        }
        if (adjustedIndex > items.size()) {
            adjustedIndex = items.size();
        }
        items.addAll(adjustedIndex, moving);
        fireTableDataChanged();
    }

    @Override
    public int getRowCount() {
        return items.size();
    }

    @Override
    public int getColumnCount() {
        return COLUMNS.length;
    }

    @Override
    public String getColumnName(int column) {
        return COLUMNS[column];
    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        FileItem item = items.get(rowIndex);
        return switch (columnIndex) {
            case 0 -> item.getFileName();
            case 1 -> item.getSize();
            case 2 -> FORMATTER.format(Instant.ofEpochMilli(item.getLastModified().toMillis()));
            default -> "";
        };
    }
}
