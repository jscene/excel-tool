package com.xz.excel;

import com.xz.excel.utils.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Fingal on 2017-5-17.
 */
public class ExcelProcessor {

    public static List<String> getHeader(Sheet sheet, int index) {
        if(sheet == null) {
            System.out.println("sheet was null...");
            return null;
        }
        Row row = sheet.getRow(index);
        if(row == null) {
            System.out.println("there has not data at "+index);
        }

        List<String> headers = new ArrayList<String>();
        for(int i=row.getFirstCellNum(); i<=row.getLastCellNum();i++) {
            if(row.getCell(i)!=null
                    && StringUtils.isNotBlank(row.getCell(i).getStringCellValue())) {
                headers.add(row.getCell(i).getStringCellValue());
            } else {
                break;
            }
        }

        return headers;
    }

    public static Integer getHeaderIndexByName(Sheet sheet, int index, String name) {
        Row header = sheet.getRow(index);
        for(int i=0; i<=header.getLastCellNum() || i<=100;i++) {
            Cell cell = header.getCell(i);
            if(cell!=null && name.equals(cell.getStringCellValue())) {
                return i;
            }
        }

        return -1;
    }

    public static void addRow(Sheet sheet, List<String> columns) {
        Row row = sheet.createRow(sheet.getLastRowNum()+1);
        for(int i=0; i<=columns.size() -1; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(columns.get(i));
        }
    }
}
