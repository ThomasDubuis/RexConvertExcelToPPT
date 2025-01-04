package org.tdubuis.config;

import com.google.gson.Gson;
import lombok.Data;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.List;

@Data
public class ConfigFile {
    private static final Logger logger = LogManager.getLogger(ConfigFile.class);


    private String excelFile;
    private String pptFile;
    private String outputFolder;
    private String excelSuffix;
    private List<Config> config;

    public String isAndReturnConfigTitle(String text) {
        for (Config c : config) {
            if (text.contains(c.title)) {
                return c.title;
            }
        }
        return null;
    }
    public Config getConfig(String title) {
        for (Config c : config) {
            if (c.title.equals(title)) {
                return c;
            }
        }
        return null;
    }

    @Data
    public static class Config {
        private String title;
        private Integer slideMonth;
        private Integer slideYTD;
        private Integer textSize;
        private Position position;
    }

    @Data
    public static class Position {
        private Integer x;
        private Integer y;
        private Integer width;
        private Integer height;
    }

    public static ConfigFile parseConfigFile(File file) {
        try {
            String fileContent = Files.readString(file.toPath());
            return new Gson().fromJson(fileContent, ConfigFile.class);
        }catch (IOException e) {
            logger.error("Error when parsing config file", e);
        }
        return null;
    }
}
