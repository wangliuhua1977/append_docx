package app.docmerge;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
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
            return mapper.readValue(configPath.toFile(), ConfigData.class);
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
    }
}
