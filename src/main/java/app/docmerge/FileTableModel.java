package app.docmerge;

import javax.swing.table.AbstractTableModel;
import java.text.DecimalFormat;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

public class FileTableModel extends AbstractTableModel {
    private static final String[] COLUMNS = {"选择", "序号", "文件名", "扩展名", "类型", "大小", "最后修改", "来源目录", "状态"};
    private static final DateTimeFormatter TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private static final DecimalFormat SIZE_FORMAT = new DecimalFormat("#,###");
    private List<FileItem> items = new ArrayList<>();

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
    public Class<?> getColumnClass(int columnIndex) {
        return switch (columnIndex) {
            case 0 -> Boolean.class;
            case 1 -> Integer.class;
            default -> String.class;
        };
    }

    @Override
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        if (columnIndex != 0) {
            return false;
        }
        FileItem item = items.get(rowIndex);
        return !item.isMissing();
    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        FileItem item = items.get(rowIndex);
        return switch (columnIndex) {
            case 0 -> item.isChecked();
            case 1 -> rowIndex + 1;
            case 2 -> item.getName();
            case 3 -> item.getExtension();
            case 4 -> item.getFileType().getLabel();
            case 5 -> formatSize(item.getSize(), item.isMissing());
            case 6 -> formatTime(item);
            case 7 -> item.getSourceDir();
            case 8 -> item.isMissing() ? "文件不存在" : "正常";
            default -> "";
        };
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        if (columnIndex == 0 && rowIndex >= 0 && rowIndex < items.size()) {
            FileItem item = items.get(rowIndex);
            if (!item.isMissing()) {
                item.setChecked(Boolean.TRUE.equals(aValue));
                fireTableCellUpdated(rowIndex, columnIndex);
            }
        }
    }

    public void setItems(List<FileItem> items) {
        this.items = items;
        fireTableDataChanged();
    }

    public List<FileItem> getItems() {
        return items;
    }

    public FileItem getItemAt(int rowIndex) {
        return items.get(rowIndex);
    }

    public void removeRows(int[] rows) {
        if (rows == null || rows.length == 0) {
            return;
        }
        for (int i = rows.length - 1; i >= 0; i--) {
            int row = rows[i];
            if (row >= 0 && row < items.size()) {
                items.remove(row);
            }
        }
        fireTableDataChanged();
    }

    public void clear() {
        items.clear();
        fireTableDataChanged();
    }

    public void setAllChecked(boolean checked) {
        for (FileItem item : items) {
            if (!item.isMissing()) {
                item.setChecked(checked);
            }
        }
        fireTableDataChanged();
    }

    public void invertChecked() {
        for (FileItem item : items) {
            if (!item.isMissing()) {
                item.setChecked(!item.isChecked());
            }
        }
        fireTableDataChanged();
    }

    private String formatSize(long size, boolean missing) {
        if (missing && size <= 0) {
            return "-";
        }
        return SIZE_FORMAT.format(size);
    }

    private String formatTime(FileItem item) {
        if (item.getLastModified() == null) {
            return "-";
        }
        LocalDateTime time = LocalDateTime.ofInstant(Instant.ofEpochMilli(item.getLastModified().toMillis()), ZoneId.systemDefault());
        return time.format(TIME_FORMATTER);
    }
}
