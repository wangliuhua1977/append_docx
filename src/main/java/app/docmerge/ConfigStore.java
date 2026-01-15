package app.docmerge;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ConfigStore {
    private final ObjectMapper mapper;
    private final Path configPath;

    public ConfigStore() {
        this.mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
        this.configPath = Path.of(System.getProperty("user.home"), ".doc-merge-app", "config.json");
    }

    public synchronized ConfigData load() {
        if (!Files.exists(configPath)) {
            return new ConfigData();
        }
        try {
            ConfigData data = mapper.readValue(configPath.toFile(), ConfigData.class);
            if (data.perDirOrder == null) {
                data.perDirOrder = new HashMap<>();
            }
            if (data.lastFileList == null) {
                data.lastFileList = new ArrayList<>();
            }
            return data;
        } catch (IOException e) {
            return new ConfigData();
        }
    }

    public synchronized void save(ConfigData data) {
        try {
            Files.createDirectories(configPath.getParent());
            mapper.writeValue(configPath.toFile(), data);
        } catch (IOException ignored) {
            // 配置写入失败时忽略
        }
    }

    public static class ConfigData {
        private String lastInputDir;
        private String lastOutputDir;
        private String lastOutputFileName;
        private Map<String, List<String>> perDirOrder = new HashMap<>();
        private List<FileEntry> lastFileList = new ArrayList<>();

        public String getLastInputDir() {
            return lastInputDir;
        }

        public void setLastInputDir(String lastInputDir) {
            this.lastInputDir = lastInputDir;
        }

        public String getLastOutputDir() {
            return lastOutputDir;
        }

        public void setLastOutputDir(String lastOutputDir) {
            this.lastOutputDir = lastOutputDir;
        }

        public String getLastOutputFileName() {
            return lastOutputFileName;
        }

        public void setLastOutputFileName(String lastOutputFileName) {
            this.lastOutputFileName = lastOutputFileName;
        }

        public Map<String, List<String>> getPerDirOrder() {
            return perDirOrder;
        }

        public void setPerDirOrder(Map<String, List<String>> perDirOrder) {
            this.perDirOrder = perDirOrder;
        }

        public List<FileEntry> getLastFileList() {
            return lastFileList;
        }

        public void setLastFileList(List<FileEntry> lastFileList) {
            this.lastFileList = lastFileList;
        }
    }

    public static class FileEntry {
        private String absolutePath;
        private boolean checked;
        private String name;
        private String extension;
        private Long size;
        private Long lastModified;
        private String sourceDir;

        public String getAbsolutePath() {
            return absolutePath;
        }

        public void setAbsolutePath(String absolutePath) {
            this.absolutePath = absolutePath;
        }

        public boolean isChecked() {
            return checked;
        }

        public void setChecked(boolean checked) {
            this.checked = checked;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getExtension() {
            return extension;
        }

        public void setExtension(String extension) {
            this.extension = extension;
        }

        public Long getSize() {
            return size;
        }

        public void setSize(Long size) {
            this.size = size;
        }

        public Long getLastModified() {
            return lastModified;
        }

        public void setLastModified(Long lastModified) {
            this.lastModified = lastModified;
        }

        public String getSourceDir() {
            return sourceDir;
        }

        public void setSourceDir(String sourceDir) {
            this.sourceDir = sourceDir;
        }
    }
}
