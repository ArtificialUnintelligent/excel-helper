package com.au.resolver;

import com.au.annotation.ExcelColumn;
import com.au.model.ExcelTitle;
import com.google.common.collect.Lists;
import org.apache.poi.hssf.usermodel.HSSFBorderFormatting;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * @author:artificialunintelligent
 * @Date:2019-07-05
 * @Time:16:30
 * @Desc: 数据量较多时建议使用分页写，按页持久化数据到磁盘，防止内存占用过多的情况导致oom的发生
 */
public class ExcelWriter {

    public void create(List<?> datas, SXSSFWorkbook wb, int page, int pageSize) throws Exception {

        if (null == datas || datas.isEmpty()) return;

        List<Field> declaredFields = Lists.newArrayList();
        Class<?> clazz = datas.get(0).getClass();
        while (null != clazz) {
            Field[] declaredField = clazz.getDeclaredFields();
            declaredFields.addAll(0, Arrays.asList(declaredField));
            clazz = clazz.getSuperclass();
        }

        List<ExcelTitle> columnFirstTitles = new ArrayList<>();
        List<String> columnSecondTitles = new ArrayList<>();
        List<Field> columnFields = new ArrayList<>();

        if (declaredFields.isEmpty()) return;

        for (Field filed : declaredFields) {
            ExcelColumn annotation = filed.getAnnotation(ExcelColumn.class);
            if (null != annotation) {
                columnFields.add(filed);
                columnSecondTitles.add(annotation.secondTitle());
                int isContained = isContained(columnFirstTitles, annotation.firstTitle());
                if (isContained == -1) {
                    ExcelTitle title = new ExcelTitle();
                    title.setTitleTxt(annotation.firstTitle());
                    title.setTitleRange(1);
                    columnFirstTitles.add(title);
                } else {
                    ExcelTitle title = columnFirstTitles.get(isContained);
                    int range = title.getTitleRange() + 1;
                    title.setTitleRange(range);
                }
            }
        }

        // cal max column charactors
        int size = columnFields.size();
        int[] columnCharacterLength = new int[size];
        for (Object next : datas) {
            for (int i = 0; i < size; i++) {
                Field columnFiled = columnFields.get(i);
                columnFiled.setAccessible(true);
                Object object = columnFiled.get(next);
                if (null == object) {
                    object = "";
                }
                if (object.toString().length() > columnCharacterLength[i]) {
                    columnCharacterLength[i] = object.toString().length();
                }
            }

        }

        Sheet sheet;
        if (wb.getSheet("sheet") == null) {
            sheet = wb.createSheet("sheet");
        } else {
            sheet = wb.getSheet("sheet");
        }

        int rowIndex = 0;
        int columnIndex = 0;
        Row row;
        Cell cell;

        int col = 0;
        if (page == 1) {
            row = sheet.createRow(rowIndex++);
            for (ExcelTitle title : columnFirstTitles) {
                int count = title.getTitleRange();
                CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, col,
                        col + count - 1);
                if (col != (col + count - 1)) {
                    sheet.addMergedRegion(cellRangeAddress);
                }
                setRegionBorder(1, cellRangeAddress, sheet, wb);
                col += count;
            }

            size = columnFirstTitles.size();
            for (int i = 0; i < size; i++) {
                cell = row.createCell(columnIndex);
                cell.setCellStyle(this.getTitleStyle(wb));
                cell.setCellValue(columnFirstTitles.get(i).getTitleTxt());
                columnIndex += columnFirstTitles.get(i).getTitleRange();
            }

            size = columnSecondTitles.size();
            row = sheet.createRow(rowIndex);
            columnIndex = 0;
            for (int i = 0; i < size; i++) {
                cell = row.createCell(columnIndex++);
                cell.setCellStyle(this.getTitleStyle(wb));
                cell.setCellValue(columnSecondTitles.get(i));
                sheet.setColumnWidth(i, calcColumnWidth(columnCharacterLength[i]));
            }
        }

        //iterator = datas.iterator();
        int total = datas.size();
        int i = 0;
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setWrapText(true);
        for (int index = pageSize * (page - 1) + 2; index < pageSize * (page - 1) + total + 2; index++) {
            Object next = datas.get(i++);
            row = sheet.createRow(index);

            columnIndex = 0;
            for (Field columnFiled : columnFields) {
                columnFiled.setAccessible(true);
                Object object = columnFiled.get(next);
                if (null == object) object = "";
                cell = row.createCell(columnIndex++);
//                cell.setCellStyle(this.getStyle(wb)); // Apply style to cell
                cell.setCellStyle(cellStyle);
                cell.setCellValue(object.toString());
            }
        }
        // sheet.autoSizeColumn(0);
    }

    private int calcColumnWidth(int charactors) {
        if (charactors < 25) return 25 * 256;
        int width = (int) (charactors * 1.14388) * 128;
        return width > 255 * 256 ? 200 * 256 : width;
    }

    private Integer isContained(List<ExcelTitle> excelTitles, String title) {
        for (int i = 0; i < excelTitles.size(); i++) {
            if (excelTitles.get(i).getTitleTxt().equals(title)) {
                return i;
            }
        }
        return -1;
    }

    private CellStyle getTitleStyle(SXSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        return this.setBorder(style);
    }

//    private CellStyle getStyle(SXSSFWorkbook workbook){
//        CellStyle style = workbook.createCellStyle();
//        style.setWrapText(true); // Set wordwrap
//        return this.setBorder(style);
//    }

    private CellStyle setBorder(CellStyle style) {
        style.setBorderTop(HSSFBorderFormatting.BORDER_THIN);
        style.setBorderBottom(HSSFBorderFormatting.BORDER_THIN);
        style.setBorderLeft(HSSFBorderFormatting.BORDER_THIN);
        style.setBorderRight(HSSFBorderFormatting.BORDER_THIN);
        style.setTopBorderColor(HSSFColor.BLACK.index);
        style.setBottomBorderColor(HSSFColor.BLACK.index);
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        style.setRightBorderColor(HSSFColor.BLACK.index);
        return style;
    }

    private void setRegionBorder(int border, CellRangeAddress region, Sheet sheet, Workbook wb) {
        RegionUtil.setBorderBottom(border, region, sheet, wb);
        RegionUtil.setBorderLeft(border, region, sheet, wb);
        RegionUtil.setBorderRight(border, region, sheet, wb);
        RegionUtil.setBorderTop(border, region, sheet, wb);
    }

}
