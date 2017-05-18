package com.xz.excel.entity;

import com.xz.excel.utils.StringUtils;

import java.util.Properties;

/**
 * Created by Fingal on 2017-5-17.
 */
public class Config {
    private String sourceFile;
    private String targetFile;
    private String[] statistics;
    private Integer headerIndex;
    private Integer contentIndex;
    public Config() {
    }

    public void load(Properties prop) {
        this.sourceFile = prop.getProperty("source_file");
        this.targetFile = prop.getProperty("target_file");
        String statisticsProp = prop.getProperty("statistics");
        this.headerIndex = Integer.valueOf(prop.getProperty("header_index", "2"));
        this.contentIndex = Integer.valueOf(prop.getProperty("content_index", "3"));

        if(StringUtils.isNotBlank(statisticsProp)){
            this.statistics = statisticsProp.split(",");
        }
    }

    public String getSourceFile() {
        return sourceFile;
    }

    public void setSourceFile(String sourceFile) {
        this.sourceFile = sourceFile;
    }

    public String getTargetFile() {
        return targetFile;
    }

    public void setTargetFile(String targetFile) {
        this.targetFile = targetFile;
    }

    public String[] getStatistics() {
        return statistics;
    }

    public void setStatistics(String[] statistics) {
        this.statistics = statistics;
    }

    public Integer getHeaderIndex() {
        return headerIndex;
    }

    public void setHeaderIndex(Integer headerIndex) {
        this.headerIndex = headerIndex;
    }

    public Integer getContentIndex() {
        return contentIndex;
    }

    public void setContentIndex(Integer contentIndex) {
        this.contentIndex = contentIndex;
    }
}
