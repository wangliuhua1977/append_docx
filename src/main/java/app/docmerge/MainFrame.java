package app.docmerge;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumnModel;
import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.io.IOException;
import java.nio.file.Path;
import java.text.DecimalFormat;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MainFrame extends JFrame {
    private final JTextField inputField = new JTextField();
    private final JTextField outputField = new JTextField();
    private final JTextField outputNameField = new JTextField();
    private final JButton selectInputButton = new JButton("选择输入目录");
    private final JButton refreshButton = new JButton("刷新文件列表");
    private final JButton selectOutputButton = new JButton("选择输出目录");
    private final JButton mergeButton = new JButton("开始合并");
    private final JButton cancelButton = new JButton("取消");
    private final JButton clearLogButton = new JButton("清空");
    private final JButton copyLogButton = new JButton("复制全部");
    private final JProgressBar progressBar = new JProgressBar();
    private final JLabel statusLabel = new JLabel("准备就绪");
    private final JTextArea logArea = new JTextArea();
    private final DefaultTableModel tableModel;
    private final JTable table;

    private final ConfigStore configStore = new ConfigStore();
    private final FileScanner fileScanner = new FileScanner();
    private final UiLogger logger = new UiLogger(logArea);
    private ConfigStore.ConfigData configData;
    private List<FileItem> currentItems = new ArrayList<>();
    private SwingWorker<Boolean, String> worker;

    public MainFrame() {
        super("文档合并工具");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1000, 700);
        setLocationRelativeTo(null);

        inputField.setEditable(false);
        outputField.setEditable(false);
        outputNameField.setText(defaultOutputName());

        String[] columns = {"序号", "文件名", "扩展名", "大小", "最后修改"};
        tableModel = new DefaultTableModel(columns, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };
        table = new JTable(tableModel);
        table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        table.setDragEnabled(true);
        table.setDropMode(javax.swing.DropMode.INSERT_ROWS);
        table.setTransferHandler(new DragReorderSupport(table, this::handleReorder));

        TableColumnModel columnModel = table.getColumnModel();
        columnModel.getColumn(0).setPreferredWidth(50);
        columnModel.getColumn(1).setPreferredWidth(400);
        columnModel.getColumn(2).setPreferredWidth(80);
        columnModel.getColumn(3).setPreferredWidth(120);
        columnModel.getColumn(4).setPreferredWidth(200);

        logArea.setEditable(false);
        logArea.setLineWrap(true);

        JPanel topPanel = buildTopPanel();
        JPanel centerPanel = buildCenterPanel();
        JPanel bottomPanel = buildBottomPanel();

        add(topPanel, BorderLayout.NORTH);
        add(centerPanel, BorderLayout.CENTER);
        add(bottomPanel, BorderLayout.SOUTH);

        bindActions();
        loadConfig();
    }

    private JPanel buildTopPanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new javax.swing.BoxLayout(panel, javax.swing.BoxLayout.Y_AXIS));

        JPanel inputPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        inputPanel.add(selectInputButton);
        inputField.setPreferredSize(new Dimension(600, 28));
        inputPanel.add(inputField);
        inputPanel.add(refreshButton);

        JPanel outputPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        outputPanel.add(selectOutputButton);
        outputField.setPreferredSize(new Dimension(600, 28));
        outputPanel.add(outputField);
        outputPanel.add(new JLabel("输出文件名"));
        outputNameField.setPreferredSize(new Dimension(220, 28));
        outputPanel.add(outputNameField);

        panel.add(inputPanel);
        panel.add(outputPanel);
        return panel;
    }

    private JPanel buildCenterPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        JScrollPane tableScroll = new JScrollPane(table);
        panel.add(tableScroll, BorderLayout.CENTER);
        return panel;
    }

    private JPanel buildBottomPanel() {
        JPanel panel = new JPanel(new BorderLayout());

        JPanel actions = new JPanel(new FlowLayout(FlowLayout.LEFT));
        actions.add(mergeButton);
        actions.add(cancelButton);
        cancelButton.setEnabled(false);

        JPanel progressPanel = new JPanel();
        progressPanel.setLayout(new javax.swing.BoxLayout(progressPanel, javax.swing.BoxLayout.Y_AXIS));
        progressBar.setStringPainted(true);
        progressBar.setString("等待开始");
        progressPanel.add(progressBar);
        progressPanel.add(statusLabel);

        JPanel logPanel = new JPanel(new BorderLayout());
        logPanel.setBorder(BorderFactory.createTitledBorder("日志"));
        JScrollPane logScroll = new JScrollPane(logArea);
        logPanel.add(logScroll, BorderLayout.CENTER);
        JPanel logActions = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        logActions.add(clearLogButton);
        logActions.add(copyLogButton);
        logPanel.add(logActions, BorderLayout.SOUTH);
        logPanel.setPreferredSize(new Dimension(800, 200));

        panel.add(actions, BorderLayout.NORTH);
        panel.add(progressPanel, BorderLayout.CENTER);
        panel.add(logPanel, BorderLayout.SOUTH);
        return panel;
    }

    private void bindActions() {
        selectInputButton.addActionListener(event -> chooseDirectory(true));
        selectOutputButton.addActionListener(event -> chooseDirectory(false));
        refreshButton.addActionListener(event -> refreshFiles());
        mergeButton.addActionListener(event -> startMerge());
        cancelButton.addActionListener(event -> cancelMerge());
        clearLogButton.addActionListener(event -> logger.clear());
        copyLogButton.addActionListener(event -> copyLog());
        outputNameField.addFocusListener(new FocusAdapter() {
            @Override
            public void focusLost(FocusEvent e) {
                configData.setLastOutputFileName(outputNameField.getText().trim());
                configStore.save(configData);
            }
        });
    }

    private void loadConfig() {
        configData = configStore.load();
        if (configData.getLastInputDir() != null) {
            inputField.setText(configData.getLastInputDir());
        }
        if (configData.getLastOutputDir() != null) {
            outputField.setText(configData.getLastOutputDir());
        }
        if (configData.getLastOutputFileName() != null) {
            outputNameField.setText(configData.getLastOutputFileName());
        }
        if (configData.getLastInputDir() != null) {
            refreshFiles();
        }
    }

    private void chooseDirectory(boolean input) {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int result = chooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            Path selected = chooser.getSelectedFile().toPath();
            if (input) {
                inputField.setText(selected.toString());
                configData.setLastInputDir(selected.toString());
                refreshFiles();
            } else {
                outputField.setText(selected.toString());
                configData.setLastOutputDir(selected.toString());
                configStore.save(configData);
            }
        }
    }

    private void refreshFiles() {
        Path inputDir = getInputDir();
        if (inputDir == null) {
            showError("请先选择输入目录");
            return;
        }
        try {
            List<FileItem> scanned = fileScanner.scan(inputDir);
            Map<String, List<String>> perDirOrder = configData.getPerDirOrder();
            List<String> savedOrder = perDirOrder.getOrDefault(inputDir.toString(), Collections.emptyList());
            List<FileItem> ordered = applySavedOrder(scanned, savedOrder);
            currentItems = ordered;
            refreshTable();
            perDirOrder.put(inputDir.toString(), ordered.stream().map(FileItem::getName).toList());
            configStore.save(configData);
            logger.info("扫描完成，共 " + ordered.size() + " 个文件");
        } catch (IOException e) {
            logger.error("扫描文件失败", e);
            showError("扫描文件失败：" + e.getMessage());
        }
    }

    private List<FileItem> applySavedOrder(List<FileItem> scanned, List<String> savedOrder) {
        if (savedOrder == null || savedOrder.isEmpty()) {
            return new ArrayList<>(scanned);
        }
        Map<String, FileItem> lookup = new HashMap<>();
        for (FileItem item : scanned) {
            lookup.put(item.getName(), item);
        }
        List<FileItem> result = new ArrayList<>();
        for (String name : savedOrder) {
            FileItem item = lookup.remove(name);
            if (item != null) {
                result.add(item);
            }
        }
        result.addAll(lookup.values().stream()
                .sorted(fileScanner.defaultComparator())
                .toList());
        return result;
    }

    private void refreshTable() {
        tableModel.setRowCount(0);
        DecimalFormat sizeFormat = new DecimalFormat("#,###");
        DateTimeFormatter timeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        int index = 1;
        for (FileItem item : currentItems) {
            String size = sizeFormat.format(item.getSize());
            LocalDateTime time = LocalDateTime.ofInstant(Instant.ofEpochMilli(item.getLastModified().toMillis()), ZoneId.systemDefault());
            tableModel.addRow(new Object[]{index++, item.getName(), item.getExtension(), size, time.format(timeFormatter)});
        }
    }

    private void handleReorder(int[] fromRows, int toRow) {
        if (fromRows == null || fromRows.length == 0) {
            return;
        }
        List<Integer> fromList = new ArrayList<>();
        for (int row : fromRows) {
            fromList.add(row);
        }
        Collections.sort(fromList);
        List<FileItem> moving = new ArrayList<>();
        for (int i = fromList.size() - 1; i >= 0; i--) {
            int row = fromList.get(i);
            moving.add(0, currentItems.remove(row));
            if (row < toRow) {
                toRow--;
            }
        }
        if (toRow > currentItems.size()) {
            toRow = currentItems.size();
        }
        currentItems.addAll(toRow, moving);
        refreshTable();
        persistOrder();
    }

    private void persistOrder() {
        Path inputDir = getInputDir();
        if (inputDir == null) {
            return;
        }
        configData.getPerDirOrder().put(inputDir.toString(), currentItems.stream().map(FileItem::getName).toList());
        configData.setLastOutputFileName(outputNameField.getText().trim());
        configStore.save(configData);
    }

    private void startMerge() {
        if (worker != null && !worker.isDone()) {
            showError("已有任务正在执行");
            return;
        }
        Path outputDir = getOutputDir();
        if (getInputDir() == null || outputDir == null) {
            showError("请选择输入和输出目录");
            return;
        }
        if (currentItems.isEmpty()) {
            showError("没有可合并的文件");
            return;
        }
        persistOrder();
        mergeButton.setEnabled(false);
        cancelButton.setEnabled(true);
        progressBar.setIndeterminate(true);
        progressBar.setString("正在准备合并");
        statusLabel.setText("开始合并");

        MergeService service = new MergeService();
        worker = new SwingWorker<>() {
            @Override
            protected Boolean doInBackground() {
                try {
                    service.merge(new ArrayList<>(currentItems), outputDir, outputNameField.getText(), logger,
                            (current, total, name) -> publish("正在处理 " + current + "/" + total + "：" + name),
                            this::isCancelled);
                    return true;
                } catch (MergeService.MergeCancelledException e) {
                    logger.warn("合并已取消：" + e.getMessage());
                    return false;
                } catch (IOException e) {
                    logger.error("合并失败", e);
                    SwingUtilities.invokeLater(() -> showError("合并失败：" + e.getMessage()));
                    return false;
                }
            }

            @Override
            protected void process(List<String> chunks) {
                if (!chunks.isEmpty()) {
                    String message = chunks.get(chunks.size() - 1);
                    progressBar.setString(message);
                    statusLabel.setText(message);
                }
            }

            @Override
            protected void done() {
                progressBar.setIndeterminate(false);
                progressBar.setValue(0);
                mergeButton.setEnabled(true);
                cancelButton.setEnabled(false);
                try {
                    Boolean success = get();
                    if (Boolean.TRUE.equals(success) && !isCancelled()) {
                        progressBar.setString("完成");
                        statusLabel.setText("完成");
                        JOptionPane.showMessageDialog(MainFrame.this, "合并完成", "完成", JOptionPane.INFORMATION_MESSAGE);
                    } else if (isCancelled()) {
                        progressBar.setString("已取消");
                        statusLabel.setText("已取消");
                    } else {
                        progressBar.setString("失败");
                        statusLabel.setText("失败");
                    }
                } catch (Exception e) {
                    progressBar.setString("失败");
                    statusLabel.setText("失败");
                }
            }
        };
        worker.execute();
    }

    private void cancelMerge() {
        if (worker != null && !worker.isDone()) {
            worker.cancel(true);
            logger.warn("用户取消了合并任务");
            progressBar.setIndeterminate(false);
            progressBar.setString("已取消");
            statusLabel.setText("已取消");
        }
    }

    private void copyLog() {
        String text = logArea.getText();
        Toolkit.getDefaultToolkit().getSystemClipboard().setContents(new StringSelection(text), null);
        logger.info("日志已复制到剪贴板");
    }

    private Path getInputDir() {
        String text = inputField.getText();
        if (text == null || text.isBlank()) {
            return null;
        }
        return Path.of(text);
    }

    private Path getOutputDir() {
        String text = outputField.getText();
        if (text == null || text.isBlank()) {
            return null;
        }
        return Path.of(text);
    }

    private String defaultOutputName() {
        return "合并结果_" + DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss").format(LocalDateTime.now()) + ".docx";
    }

    private void showError(String message) {
        JOptionPane.showMessageDialog(this, message, "错误", JOptionPane.ERROR_MESSAGE);
    }
}
