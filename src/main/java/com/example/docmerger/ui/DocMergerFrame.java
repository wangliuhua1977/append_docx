package com.example.docmerger.ui;

import com.example.docmerger.model.FileItem;
import com.example.docmerger.model.FileTableModel;
import com.example.docmerger.service.FileScanner;
import com.example.docmerger.service.OrderPersistence;
import com.example.docmerger.util.NaturalSortComparator;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.table.TableRowSorter;
import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

public final class DocMergerFrame extends JFrame {
    private static final Logger LOGGER = Logger.getLogger(DocMergerFrame.class.getName());

    private final JTextField sourceDirField = new JTextField(30);
    private final JButton browseButton = new JButton("Browse");
    private final JButton rescanButton = new JButton("Rescan");
    private final JButton resetOrderButton = new JButton("Reset to Default Order");
    private final JButton mergeButton = new JButton("Start Merge");
    private final JLabel orderStatusLabel = new JLabel("Order: Default");
    private final JTextArea logArea = new JTextArea(10, 80);

    private final FileTableModel tableModel = new FileTableModel();
    private final JTable table = new JTable(tableModel);
    private final AtomicBoolean dragEnabled = new AtomicBoolean(true);
    private final OrderPersistence orderPersistence = new OrderPersistence(DocMergerFrame.class);

    private List<FileItem> defaultItems = new ArrayList<>();
    private boolean customOrderActive = false;

    public DocMergerFrame() {
        super("Doc Merger");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setMinimumSize(new Dimension(900, 600));

        logArea.setEditable(false);
        logArea.setLineWrap(true);
        logArea.setWrapStyleWord(true);

        configureTable();
        layoutComponents();
        bindActions();
        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                if (customOrderActive) {
                    persistCustomOrder();
                } else {
                    orderPersistence.clearCustomOrder();
                }
            }
        });
    }

    private void configureTable() {
        table.setFillsViewportHeight(true);
        table.setDragEnabled(true);
        table.setDropMode(javax.swing.DropMode.INSERT_ROWS);
        table.setTransferHandler(new ReorderableTableTransferHandler(table, dragEnabled, this::handleRowsMoved));

        TableRowSorter<FileTableModel> sorter = new TableRowSorter<>(tableModel);
        sorter.setSortable(0, false);
        sorter.setSortable(1, false);
        sorter.setSortable(2, false);
        table.setRowSorter(sorter);
    }

    private void layoutComponents() {
        JPanel topPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(4, 4, 4, 4);
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.WEST;
        topPanel.add(new JLabel("Source Directory:"), gbc);

        gbc.gridx = 1;
        gbc.weightx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        topPanel.add(sourceDirField, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        gbc.fill = GridBagConstraints.NONE;
        topPanel.add(browseButton, gbc);

        gbc.gridx = 3;
        topPanel.add(rescanButton, gbc);

        gbc.gridx = 0;
        gbc.gridy = 1;
        topPanel.add(orderStatusLabel, gbc);

        gbc.gridx = 1;
        topPanel.add(resetOrderButton, gbc);

        JPanel centerPanel = new JPanel(new BorderLayout());
        centerPanel.setBorder(BorderFactory.createTitledBorder("Files"));
        centerPanel.add(new JScrollPane(table), BorderLayout.CENTER);

        JPanel bottomPanel = new JPanel(new BorderLayout());
        JPanel mergePanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        mergePanel.add(mergeButton);
        bottomPanel.add(mergePanel, BorderLayout.NORTH);
        bottomPanel.add(new JScrollPane(logArea), BorderLayout.CENTER);

        add(topPanel, BorderLayout.NORTH);
        add(centerPanel, BorderLayout.CENTER);
        add(bottomPanel, BorderLayout.SOUTH);
    }

    private void bindActions() {
        browseButton.addActionListener(event -> chooseDirectory());
        rescanButton.addActionListener(event -> rescan());
        resetOrderButton.addActionListener(event -> resetToDefaultOrder());
        mergeButton.addActionListener(event -> startMerge());
    }

    private void chooseDirectory() {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int result = chooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            sourceDirField.setText(chooser.getSelectedFile().getAbsolutePath());
            rescan();
        }
    }

    private void rescan() {
        Path directory = resolveSourceDirectory();
        if (directory == null) {
            return;
        }
        rescanButton.setEnabled(false);
        logInfo("Rescanning directory: " + directory);
        SwingWorker<List<FileItem>, Void> worker = new SwingWorker<>() {
            @Override
            protected List<FileItem> doInBackground() throws Exception {
                FileScanner scanner = new FileScanner();
                List<FileItem> items = scanner.scan(directory);
                items.sort(new NaturalSortComparator());
                return items;
            }

            @Override
            protected void done() {
                rescanButton.setEnabled(true);
                try {
                    List<FileItem> items = get();
                    defaultItems = new ArrayList<>(items);
                    applyCustomOrder(directory, items);
                } catch (Exception ex) {
                    logError("Rescan failed: " + ex.getMessage(), ex);
                }
            }
        };
        worker.execute();
    }

    private void applyCustomOrder(Path directory, List<FileItem> scannedItems) {
        Optional<OrderPersistence.StoredOrder> stored = orderPersistence.loadCustomOrder();
        if (stored.isEmpty()) {
            tableModel.setItems(scannedItems);
            setCustomOrderActive(false);
            logInfo("Custom order not applied: no custom order stored.");
            return;
        }
        OrderPersistence.StoredOrder storedOrder = stored.get();
        if (!storedOrder.sourceDir().equals(directory.toString())) {
            tableModel.setItems(scannedItems);
            setCustomOrderActive(false);
            logInfo("Custom order not applied: source directory mismatch.");
            return;
        }

        Map<String, FileItem> byStableId = new HashMap<>();
        for (FileItem item : scannedItems) {
            byStableId.put(normalizeStableId(item.getStableId()), item);
        }

        List<FileItem> aligned = new ArrayList<>();
        int missingCount = 0;
        for (String stableId : storedOrder.stableIds()) {
            FileItem item = byStableId.remove(normalizeStableId(stableId));
            if (item != null) {
                aligned.add(item);
            } else {
                missingCount++;
            }
        }
        int newCount = byStableId.size();
        if (!byStableId.isEmpty()) {
            List<FileItem> newItems = new ArrayList<>(byStableId.values());
            newItems.sort(new NaturalSortComparator());
            aligned.addAll(newItems);
        }
        tableModel.setItems(aligned);
        setCustomOrderActive(true);
        logInfo(String.format("Custom order restored: %d items; missing %d items; new %d items appended.",
                aligned.size() - newCount, missingCount, newCount));
    }

    private void handleRowsMoved(int[] rows, int dropIndex) {
        if (rows.length == 0) {
            return;
        }
        int originalIndex = rows[0];
        int targetIndex = dropIndex;
        if (originalIndex < dropIndex) {
            targetIndex = dropIndex - rows.length;
        }
        setCustomOrderActive(true);
        if (rows.length == 1 && targetIndex >= 0 && targetIndex < tableModel.getRowCount()) {
            String fileName = tableModel.getItems().get(targetIndex).getFileName();
            logInfo(String.format("Moved \"%s\" from %d to %d", fileName, originalIndex + 1, targetIndex + 1));
        } else {
            logInfo(String.format("Moved %d rows starting at %d to %d", rows.length, originalIndex + 1, targetIndex + 1));
        }
        persistCustomOrder();
    }

    private void persistCustomOrder() {
        List<String> ids = tableModel.getItems().stream()
                .map(FileItem::getStableId)
                .toList();
        String sourceDir = sourceDirField.getText().trim();
        if (!sourceDir.isEmpty()) {
            orderPersistence.saveCustomOrder(sourceDir, ids);
        }
    }

    private void resetToDefaultOrder() {
        tableModel.setItems(defaultItems);
        orderPersistence.clearCustomOrder();
        setCustomOrderActive(false);
        logInfo("Reset to default order.");
    }

    private void startMerge() {
        if (tableModel.getRowCount() == 0) {
            JOptionPane.showMessageDialog(this, "No files to merge.");
            return;
        }
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Select output file");
        int result = chooser.showSaveDialog(this);
        if (result != JFileChooser.APPROVE_OPTION) {
            return;
        }
        Path output = chooser.getSelectedFile().toPath();
        dragEnabled.set(false);
        table.setDragEnabled(false);
        mergeButton.setEnabled(false);
        logInfo("Starting merge to: " + output);
        Instant start = Instant.now();
        SwingWorker<Void, Void> worker = new SwingWorker<>() {
            @Override
            protected Void doInBackground() throws Exception {
                mergeFiles(output, tableModel.getItems());
                return null;
            }

            @Override
            protected void done() {
                dragEnabled.set(true);
                table.setDragEnabled(true);
                mergeButton.setEnabled(true);
                try {
                    get();
                    Duration elapsed = Duration.between(start, Instant.now());
                    logInfo("Merge completed in " + elapsed.toSeconds() + "s.");
                } catch (Exception ex) {
                    logError("Merge failed: " + ex.getMessage(), ex);
                }
            }
        };
        worker.execute();
    }

    private void mergeFiles(Path output, List<FileItem> items) throws IOException {
        if (Files.exists(output)) {
            Files.delete(output);
        }
        try (var outputStream = Files.newOutputStream(output)) {
            for (FileItem item : items) {
                Files.copy(item.getPath(), outputStream);
            }
        }
    }

    private Path resolveSourceDirectory() {
        String text = sourceDirField.getText().trim();
        if (text.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please select a source directory.");
            return null;
        }
        Path directory = Path.of(text);
        if (!Files.isDirectory(directory)) {
            JOptionPane.showMessageDialog(this, "Source directory does not exist.");
            return null;
        }
        return directory;
    }

    private void setCustomOrderActive(boolean active) {
        customOrderActive = active;
        orderStatusLabel.setText(active ? "Order: Custom" : "Order: Default");
    }

    private String normalizeStableId(String stableId) {
        return stableId.toLowerCase(Locale.ROOT);
    }

    private void logInfo(String message) {
        LOGGER.info(message);
        SwingUtilities.invokeLater(() -> logArea.append(message + "\n"));
    }

    private void logError(String message, Exception ex) {
        LOGGER.log(Level.SEVERE, message, ex);
        SwingUtilities.invokeLater(() -> logArea.append(message + "\n"));
    }
}
