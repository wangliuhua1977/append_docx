package app.docmerge;

import javax.swing.BorderFactory;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.TableColumnModel;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.BasicFileAttributes;
import java.nio.file.attribute.FileTime;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

public class MainFrame extends JFrame {
    private static final Pattern ILLEGAL_NAME = Pattern.compile("[\\\\/:*?\"<>|]");

    private final JTextField inputField = new JTextField();
    private final JTextField outputField = new JTextField();
    private final JTextField outputNameField = new JTextField();
    private final JButton selectInputButton = new JButton("选择输入目录");
    private final JButton refreshButton = new JButton("刷新文件列表");
    private final JButton addFilesButton = new JButton("添加文件...");
    private final JButton addDirButton = new JButton("添加目录...");
    private final JButton removeSelectedButton = new JButton("移除选中");
    private final JButton clearListButton = new JButton("清空列表");
    private final JButton selectAllButton = new JButton("全选");
    private final JButton selectNoneButton = new JButton("全不选");
    private final JButton invertSelectButton = new JButton("反选");
    private final JButton selectOutputButton = new JButton("选择输出目录");
    private final JButton mergeButton = new JButton("开始合并");
    private final JButton cancelButton = new JButton("取消");
    private final JRadioButton autoModeButton = new JRadioButton("自动（Word优先）");
    private final JRadioButton wordOnlyButton = new JRadioButton("仅 Word");
    private final JRadioButton wpsOnlyButton = new JRadioButton("仅 WPS");
    private final JButton probeEnvButton = new JButton("检测环境");
    private final JLabel wordStatusLabel = new JLabel();
    private final JLabel wpsStatusLabel = new JLabel();
    private final JLabel modeStatusLabel = new JLabel();
    private final JButton clearLogButton = new JButton("清空");
    private final JButton copyLogButton = new JButton("复制全部");
    private final JProgressBar progressBar = new JProgressBar();
    private final JLabel statusLabel = new JLabel("准备就绪");
    private final JTextArea logArea = new JTextArea();
    private final FileTableModel tableModel = new FileTableModel();
    private final JTable table = new JTable(tableModel);

    private final ConfigStore configStore = new ConfigStore();
    private final FileScanner fileScanner = new FileScanner();
    private final UiLogger logger = new UiLogger(logArea);
    private final DocComConverterSelector converterSelector = new DocComConverterSelector();
    private final DocComConverterResolver converterResolver = new DocComConverterResolver(converterSelector);
    private ConfigStore.ConfigData configData;
    private List<FileItem> currentItems = new ArrayList<>();
    private SwingWorker<Boolean, String> worker;
    private DocComConverterResolver.Resolution probeResolution;

    public MainFrame() {
        super("Word文档合并工具-- 黄莉专用版");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1200, 760);
        setLocationRelativeTo(null);

        inputField.setEditable(false);
        outputField.setEditable(false);
        outputNameField.setText(defaultOutputName());

        table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        table.setDragEnabled(true);
        table.setDropMode(javax.swing.DropMode.INSERT_ROWS);
        table.setTransferHandler(new DragReorderSupport(table, this::handleReorder));
        table.setDefaultRenderer(Object.class, new MissingAwareRenderer(tableModel));

        TableColumnModel columnModel = table.getColumnModel();
        columnModel.getColumn(0).setPreferredWidth(60);
        columnModel.getColumn(1).setPreferredWidth(60);
        columnModel.getColumn(2).setPreferredWidth(320);
        columnModel.getColumn(3).setPreferredWidth(80);
        columnModel.getColumn(4).setPreferredWidth(100);
        columnModel.getColumn(5).setPreferredWidth(160);
        columnModel.getColumn(6).setPreferredWidth(240);
        columnModel.getColumn(7).setPreferredWidth(100);

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
        refreshComProbeStatus(true);
        logComProbeStatus();
    }

    private JPanel buildTopPanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new javax.swing.BoxLayout(panel, javax.swing.BoxLayout.Y_AXIS));

        JPanel inputPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        inputPanel.add(selectInputButton);
        inputField.setPreferredSize(new Dimension(550, 28));
        inputPanel.add(inputField);
        inputPanel.add(refreshButton);

        JPanel listActions = new JPanel(new FlowLayout(FlowLayout.LEFT));
        listActions.add(addFilesButton);
        listActions.add(addDirButton);
        listActions.add(removeSelectedButton);
        listActions.add(clearListButton);
        listActions.add(selectAllButton);
        listActions.add(selectNoneButton);
        listActions.add(invertSelectButton);

        JPanel outputPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        outputPanel.add(selectOutputButton);
        outputField.setPreferredSize(new Dimension(550, 28));
        outputPanel.add(outputField);
        outputPanel.add(new JLabel("输出文件名"));
        outputNameField.setPreferredSize(new Dimension(240, 28));
        outputPanel.add(outputNameField);

        JPanel docEnginePanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        docEnginePanel.add(new JLabel("DOC 转换引擎"));
        ButtonGroup group = new ButtonGroup();
        group.add(autoModeButton);
        group.add(wordOnlyButton);
        group.add(wpsOnlyButton);
        docEnginePanel.add(autoModeButton);
        docEnginePanel.add(wordOnlyButton);
        docEnginePanel.add(wpsOnlyButton);
        docEnginePanel.add(probeEnvButton);

        JPanel statusPanel = new JPanel();
        statusPanel.setLayout(new javax.swing.BoxLayout(statusPanel, javax.swing.BoxLayout.Y_AXIS));
        statusPanel.add(wordStatusLabel);
        statusPanel.add(wpsStatusLabel);
        statusPanel.add(modeStatusLabel);
        docEnginePanel.add(statusPanel);

        panel.add(inputPanel);
        panel.add(listActions);
        panel.add(outputPanel);
        panel.add(docEnginePanel);
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
        logPanel.setPreferredSize(new Dimension(800, 220));

        panel.add(actions, BorderLayout.NORTH);
        panel.add(progressPanel, BorderLayout.CENTER);
        panel.add(logPanel, BorderLayout.SOUTH);
        return panel;
    }

    private void bindActions() {
        selectInputButton.addActionListener(event -> chooseDirectory(true));
        selectOutputButton.addActionListener(event -> chooseDirectory(false));
        refreshButton.addActionListener(event -> refreshFiles());
        addFilesButton.addActionListener(event -> addFiles());
        addDirButton.addActionListener(event -> addDirectory());
        removeSelectedButton.addActionListener(event -> removeSelected());
        clearListButton.addActionListener(event -> clearList());
        selectAllButton.addActionListener(event -> selectAll());
        selectNoneButton.addActionListener(event -> selectNone());
        invertSelectButton.addActionListener(event -> invertSelection());
        mergeButton.addActionListener(event -> startMerge());
        cancelButton.addActionListener(event -> cancelMerge());
        clearLogButton.addActionListener(event -> logger.clear());
        copyLogButton.addActionListener(event -> copyLog());
        probeEnvButton.addActionListener(event -> {
            refreshComProbeStatus(true);
            logComProbeStatus();
        });
        autoModeButton.addActionListener(event -> handleModeChange());
        wordOnlyButton.addActionListener(event -> handleModeChange());
        wpsOnlyButton.addActionListener(event -> handleModeChange());
        outputNameField.addFocusListener(new FocusAdapter() {
            @Override
            public void focusLost(FocusEvent e) {
                persistState();
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
        if (configData.getLastOutputFileName() != null && !configData.getLastOutputFileName().isBlank()) {
            outputNameField.setText(configData.getLastOutputFileName());
        }
        DocConverterMode mode = DocConverterMode.fromConfig(configData.getDocConverterMode());
        applyModeSelection(mode);
        if (!configData.getLastFileList().isEmpty()) {
            currentItems = loadFileList(configData.getLastFileList());
            refreshTable();
        } else if (configData.getLastInputDir() != null) {
            refreshFiles();
        }
    }

    private List<FileItem> loadFileList(List<ConfigStore.FileEntry> entries) {
        List<FileItem> items = new ArrayList<>();
        for (ConfigStore.FileEntry entry : entries) {
            if (entry.getAbsolutePath() == null) {
                continue;
            }
            Path path = Path.of(entry.getAbsolutePath());
            if (Files.exists(path)) {
                try {
                    BasicFileAttributes attrs = Files.readAttributes(path, BasicFileAttributes.class);
                    String sourceDir = entry.getSourceDir();
                    if (sourceDir == null || sourceDir.isBlank()) {
                        sourceDir = path.getParent() == null ? "" : path.getParent().toString();
                    }
                    FileItem item = new FileItem(path,
                            path.getFileName().toString(),
                            fileScanner.extensionOf(path.getFileName().toString()),
                            attrs.size(),
                            attrs.lastModifiedTime(),
                            sourceDir,
                            entry.isChecked(),
                            FileItem.Status.OK);
                    if (isDocBlocked(item)) {
                        continue;
                    }
                    items.add(item);
                } catch (IOException e) {
                    items.add(createMissingItem(entry, path));
                }
            } else {
                FileItem missing = createMissingItem(entry, path);
                if (!isDocBlocked(missing)) {
                    items.add(missing);
                    logger.warn("文件不存在，已标记为缺失：" + path);
                }
            }
        }
        return items;
    }

    private FileItem createMissingItem(ConfigStore.FileEntry entry, Path path) {
        String name = entry.getName();
        if (name == null || name.isBlank()) {
            name = path.getFileName() == null ? path.toString() : path.getFileName().toString();
        }
        String extension = entry.getExtension();
        if (extension == null || extension.isBlank()) {
            extension = fileScanner.extensionOf(name);
        }
        long size = entry.getSize() == null ? 0L : entry.getSize();
        FileTime lastModified = entry.getLastModified() == null ? null : FileTime.fromMillis(entry.getLastModified());
        String sourceDir = entry.getSourceDir();
        if (sourceDir == null || sourceDir.isBlank()) {
            sourceDir = path.getParent() == null ? "" : path.getParent().toString();
        }
        return new FileItem(path, name, extension, size, lastModified, sourceDir, false, FileItem.Status.MISSING);
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
                persistState();
            }
        }
    }

    private void refreshFiles() {
        Path inputDir = getInputDir();
        if (inputDir == null) {
            showError("请先选择输入目录");
            return;
        }
        addDirectoryFiles(inputDir, true);
    }

    private void addDirectory() {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int result = chooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            addDirectoryFiles(chooser.getSelectedFile().toPath(), false);
        }
    }

    private void addDirectoryFiles(Path directory, boolean replaceSameDir) {
        try {
            List<FileItem> scanned = fileScanner.scan(directory);
            scanned = filterDocFilesIfUnsupported(scanned, "目录扫描");
            Map<String, List<String>> perDirOrder = configData.getPerDirOrder();
            List<String> savedOrder = perDirOrder.getOrDefault(directory.toString(), Collections.emptyList());
            List<FileItem> ordered = applySavedOrder(scanned, savedOrder);
            if (replaceSameDir) {
                currentItems.removeIf(item -> directory.toString().equals(item.getSourceDir()));
            }
            int added = addItems(ordered);
            perDirOrder.put(directory.toString(), ordered.stream().map(FileItem::getName).toList());
            persistState();
            logger.info("扫描完成，共添加 " + added + " 个文件（目录：" + directory + "）");
        } catch (IOException e) {
            logger.error("扫描文件失败", e);
            showError("扫描文件失败：" + e.getMessage());
        }
    }

    private void addFiles() {
        JFileChooser chooser = new JFileChooser();
        chooser.setMultiSelectionEnabled(true);
        chooser.setFileFilter(new FileNameExtensionFilter("Word 文档 (*.doc, *.docx)", "doc", "docx"));
        int result = chooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            var files = chooser.getSelectedFiles();
            List<FileItem> items = new ArrayList<>();
            for (var file : files) {
                fileScanner.createFileItem(file.toPath(), true).ifPresent(items::add);
            }
            items = filterDocFilesIfUnsupported(items, "添加文件");
            items.sort(fileScanner.defaultComparator());
            int added = addItems(items);
            persistState();
            logger.info("已添加文件 " + added + " 个");
        }
    }

    private int addItems(List<FileItem> items) {
        Set<String> existing = new HashSet<>();
        for (FileItem item : currentItems) {
            existing.add(item.getPath().toAbsolutePath().toString());
        }
        int added = 0;
        for (FileItem item : items) {
            String key = item.getPath().toAbsolutePath().toString();
            if (existing.contains(key)) {
                logger.warn("已存在，跳过：" + item.getPath());
                continue;
            }
            currentItems.add(item);
            existing.add(key);
            added++;
        }
        refreshTable();
        return added;
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
        tableModel.setItems(currentItems);
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
            if (row >= 0 && row < currentItems.size()) {
                moving.add(0, currentItems.remove(row));
                if (row < toRow) {
                    toRow--;
                }
            }
        }
        if (toRow > currentItems.size()) {
            toRow = currentItems.size();
        }
        currentItems.addAll(toRow, moving);
        refreshTable();
        persistState();
    }

    private void removeSelected() {
        int[] rows = table.getSelectedRows();
        if (rows.length == 0) {
            showError("请选择需要移除的行");
            return;
        }
        tableModel.removeRows(rows);
        currentItems = tableModel.getItems();
        persistState();
        logger.info("已移除选中项：" + rows.length + " 行");
    }

    private void clearList() {
        int confirm = JOptionPane.showConfirmDialog(this, "确认清空列表吗？", "确认", JOptionPane.YES_NO_OPTION);
        if (confirm != JOptionPane.YES_OPTION) {
            return;
        }
        tableModel.clear();
        currentItems = tableModel.getItems();
        persistState();
        logger.info("列表已清空");
    }

    private void selectAll() {
        tableModel.setAllChecked(true);
        persistState();
    }

    private void selectNone() {
        tableModel.setAllChecked(false);
        persistState();
    }

    private void invertSelection() {
        tableModel.invertChecked();
        persistState();
    }

    private void persistState() {
        configData.setLastOutputFileName(outputNameField.getText().trim());
        configData.setLastFileList(serializeItems());
        configData.setPerDirOrder(buildPerDirOrder());
        configData.setDocConverterMode(getSelectedMode().name());
        configStore.save(configData);
    }

    private Map<String, List<String>> buildPerDirOrder() {
        Map<String, List<String>> result = new HashMap<>();
        for (FileItem item : currentItems) {
            String source = item.getSourceDir();
            if (source == null) {
                source = "";
            }
            result.computeIfAbsent(source, key -> new ArrayList<>()).add(item.getName());
        }
        return result;
    }

    private List<ConfigStore.FileEntry> serializeItems() {
        List<ConfigStore.FileEntry> entries = new ArrayList<>();
        for (FileItem item : currentItems) {
            ConfigStore.FileEntry entry = new ConfigStore.FileEntry();
            entry.setAbsolutePath(item.getPath().toAbsolutePath().toString());
            entry.setChecked(item.isChecked());
            entry.setName(item.getName());
            entry.setExtension(item.getExtension());
            entry.setSize(item.getSize());
            entry.setLastModified(item.getLastModified() == null ? null : item.getLastModified().toMillis());
            entry.setSourceDir(item.getSourceDir());
            entries.add(entry);
        }
        return entries;
    }

    private void startMerge() {
        if (worker != null && !worker.isDone()) {
            showError("已有任务正在执行");
            return;
        }
        Path outputDir = getOutputDir();
        if (outputDir == null) {
            showError("请选择输出目录");
            return;
        }
        if (currentItems.isEmpty()) {
            showError("没有可合并的文件");
            return;
        }
        String outputName = normalizeOutputName();
        if (outputName == null) {
            return;
        }
        Path outputFile = outputDir.resolve(outputName);
        if (Files.exists(outputFile)) {
            int overwrite = JOptionPane.showConfirmDialog(this, "输出文件已存在，是否覆盖？", "确认覆盖", JOptionPane.YES_NO_OPTION);
            if (overwrite != JOptionPane.YES_OPTION) {
                return;
            }
        }

        List<FileItem> toMerge = new ArrayList<>();
        for (FileItem item : currentItems) {
            if (item.isMissing()) {
                logger.warn("文件不存在，已跳过：" + item.getPath());
                continue;
            }
            if (item.isChecked()) {
                toMerge.add(item);
            }
        }
        refreshComProbeStatus(false);
        if (toMerge.stream().anyMatch(this::isDocItem) && !isDocConversionAvailable()) {
            DocComConverterResolver.Resolution resolution = resolveProbe(false);
            String message = resolution.errorMessage() == null
                    ? "当前环境无法进行 .doc 完美转换。请移除 .doc 文件后重试。"
                    : resolution.errorMessage();
            showError(message);
            logProbeFailure(resolution);
            return;
        }
        if (toMerge.isEmpty()) {
            showError("未选择任何待合并文件");
            return;
        }
        configData.setLastOutputFileName(outputName);
        persistState();
        mergeButton.setEnabled(false);
        cancelButton.setEnabled(true);
        progressBar.setIndeterminate(true);
        progressBar.setString("正在准备合并");
        statusLabel.setText("开始合并");

        MergeService service = new MergeService();
        DocConverterMode mode = getSelectedMode();
        worker = new SwingWorker<>() {
            @Override
            protected Boolean doInBackground() {
                try {
                    service.merge(new ArrayList<>(toMerge), outputDir, outputName, mode, converterResolver, logger,
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

    private String normalizeOutputName() {
        String text = outputNameField.getText();
        if (text == null || text.trim().isBlank()) {
            showError("请输入输出文件名");
            return null;
        }
        String name = text.trim();
        if (ILLEGAL_NAME.matcher(name).find()) {
            showError("输出文件名包含非法字符：\\ / : * ? \" < > |");
            return null;
        }
        if (!name.toLowerCase(Locale.ROOT).endsWith(".docx")) {
            name = name + ".docx";
        }
        outputNameField.setText(name);
        return name;
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

    private List<FileItem> filterDocFilesIfUnsupported(List<FileItem> items, String action) {
        if (isDocConversionAvailable()) {
            return items;
        }
        List<FileItem> filtered = new ArrayList<>();
        List<FileItem> blocked = new ArrayList<>();
        for (FileItem item : items) {
            if (isDocItem(item)) {
                blocked.add(item);
            } else {
                filtered.add(item);
            }
        }
        if (!blocked.isEmpty()) {
            DocComConverterResolver.Resolution resolution = resolveProbe(false);
            String reason = resolution.errorMessage() == null
                    ? "当前环境无法进行 .doc 完美转换"
                    : resolution.errorMessage();
            String message = reason + "。已阻止加入 .doc 文件（" + action + "）：共 " + blocked.size() + " 个";
            logger.warn(message);
            logProbeFailure(resolution);
            showError(message);
        }
        return filtered;
    }

    private boolean isDocBlocked(FileItem item) {
        if (!isDocConversionAvailable() && isDocItem(item)) {
            DocComConverterResolver.Resolution resolution = resolveProbe(false);
            String reason = resolution.errorMessage() == null
                    ? "当前环境无法进行 .doc 完美转换"
                    : resolution.errorMessage();
            logger.warn(reason + "。已跳过：" + item.getPath());
            logProbeFailure(resolution);
            return true;
        }
        return false;
    }

    private boolean isDocItem(FileItem item) {
        return "doc".equalsIgnoreCase(item.getExtension());
    }

    private void refreshComProbeStatus(boolean forceRefresh) {
        DocComConverterResolver.Resolution resolution = resolveProbe(forceRefresh);
        DocComConverterSelector.ProbeSummary summary = resolution.probeSummary();
        String wordStatus = summary.word().available() ? "可用" : "不可用";
        String wpsStatus = summary.wps().available() ? "可用" : "不可用";
        wordStatusLabel.setText("Word： " + wordStatus);
        wpsStatusLabel.setText("WPS： " + wpsStatus);
        modeStatusLabel.setText("当前模式： " + resolution.modeLabel());
    }

    private void logComProbeStatus() {
        if (probeResolution == null) {
            return;
        }
        DocComConverterSelector.ProbeSummary summary = probeResolution.probeSummary();
        logger.info("Word 完美转换：" + (summary.word().available() ? "可用" : "不可用"));
        logger.info("WPS 完美转换：" + (summary.wps().available() ? "可用" : "不可用"));
        logger.info("DOC 转换模式：" + probeResolution.modeLabel());
    }

    private boolean isDocConversionAvailable() {
        DocComConverterResolver.Resolution resolution = resolveProbe(false);
        return resolution.selection() != null;
    }

    private void logProbeFailure(DocComConverterResolver.Resolution resolution) {
        if (resolution == null || resolution.probeSummary() == null) {
            logger.warn("COM 探测失败：未知原因");
            return;
        }
        DocComConverterSelector.EngineStatus word = resolution.probeSummary().word();
        DocComConverterSelector.EngineStatus wps = resolution.probeSummary().wps();
        logger.warn("Word 探测状态：" + word.message() + "，exitCode=" + word.exitCode()
                + "\nstderr:\n" + word.stderr());
        logger.warn("WPS 探测状态：" + wps.message() + "，exitCode=" + wps.exitCode()
                + "\nstderr:\n" + wps.stderr());
    }

    private DocComConverterResolver.Resolution resolveProbe(boolean forceRefresh) {
        DocConverterMode mode = getSelectedMode();
        probeResolution = converterResolver.resolve(mode, forceRefresh);
        return probeResolution;
    }

    private DocConverterMode getSelectedMode() {
        if (wordOnlyButton.isSelected()) {
            return DocConverterMode.WORD_ONLY;
        }
        if (wpsOnlyButton.isSelected()) {
            return DocConverterMode.WPS_ONLY;
        }
        return DocConverterMode.AUTO;
    }

    private void applyModeSelection(DocConverterMode mode) {
        switch (mode) {
            case WORD_ONLY -> wordOnlyButton.setSelected(true);
            case WPS_ONLY -> wpsOnlyButton.setSelected(true);
            default -> autoModeButton.setSelected(true);
        }
    }

    private void handleModeChange() {
        configData.setDocConverterMode(getSelectedMode().name());
        persistState();
        refreshComProbeStatus(false);
    }

    private static class MissingAwareRenderer extends DefaultTableCellRenderer {
        private final FileTableModel model;

        public MissingAwareRenderer(FileTableModel model) {
            this.model = model;
        }

        @Override
        public java.awt.Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
            java.awt.Component comp = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
            if (row >= 0 && row < model.getRowCount()) {
                FileItem item = model.getItemAt(row);
                if (item.isMissing()) {
                    comp.setForeground(Color.GRAY);
                } else {
                    comp.setForeground(isSelected ? table.getSelectionForeground() : table.getForeground());
                }
            }
            return comp;
        }
    }
}
