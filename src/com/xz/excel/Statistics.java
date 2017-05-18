package com.xz.excel;

import com.xz.excel.entity.Config;
import com.xz.excel.utils.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;


/**
 * Created by Fingal on 2017-5-17.
 */
public class Statistics {
    public void process(Config config) {
        try {
            //1. read data from excel
            System.out.println("开始读取文件");
            Workbook sourceBook = createWorkbook(config.getSourceFile());
            if(sourceBook == null) {
                System.out.println("读取文件失败....");
                return ;
            }
            System.out.println("开始统计文件");
            //2. statistic data
            Workbook targetBook  = process(sourceBook, config);
            //3. write result to excel
            System.out.println("输出结果");
            targetBook.write(new FileOutputStream(config.getTargetFile()));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private Workbook process(Workbook sourceBook, Config config ){
        Workbook targetBook = new XSSFWorkbook();
        for(int i=0; i<= sourceBook.getNumberOfSheets()-1;i++) {
            Sheet sheet = sourceBook.getSheetAt(i);
            String sheetName = sheet.getSheetName();
            if(config.getStatistics() == null ) {
                break;
            }

            Sheet targetSheet = targetBook.createSheet(sheetName + "统计结果");
            for(String column : config.getStatistics()) {
                int columnIndex = ExcelProcessor.getHeaderIndexByName(sheet, config.getHeaderIndex(), column);
                if(columnIndex >= 0) {
                    writeResultHeader(targetSheet, new String[]{column, "件数"});
                    writeResult(targetSheet, processOneColumn(sheet, config.getContentIndex(), columnIndex));
                    ExcelProcessor.addRow(targetSheet, Arrays.asList(new String[]{"", ""}));
                }
            }
        }
        return targetBook;
    }

    private Map<String, Integer> processOneColumn(Sheet sheet, Integer contentIndex, Integer columnIndex) {
        Map<String, Integer> result = new HashMap<String, Integer>();
        for(int i=contentIndex; i<=sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if(row == null) { continue; }
            Cell cell = row.getCell(columnIndex);
            if(cell == null){ continue;}
            String value = cell.getStringCellValue();
            if(StringUtils.isNotBlank(value)) {
                Integer count = result.getOrDefault(value, 0);
                result.put(value, ++count);
            }
        }
        return result;
    }

    private void writeResult(Sheet sheet, Map<String, Integer> data) {
        if(data == null) {
            System.out.println("result is empty....sheetName:" + sheet.getSheetName());
            return;
        }
        for(String key : data.keySet()) {
            ExcelProcessor.addRow(sheet, Arrays.asList(new String[]{key, String.valueOf(data.get(key))}));
        }
    }

    private void writeResultHeader(Sheet sheet, String[] headers) {
        ExcelProcessor.addRow(sheet, Arrays.asList(headers));
    }

    private Workbook createWorkbook(String file) {
        Workbook workbook = null;
        try {
            workbook = new XSSFWorkbook(new  FileInputStream(file));
        } catch (Exception e) {
            e.printStackTrace();
            try {
                workbook = new HSSFWorkbook(new FileInputStream(file));
            } catch (IOException e1) {
                e1.printStackTrace();
            }
        }

        return workbook;
    }

}
