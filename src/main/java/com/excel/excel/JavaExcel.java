package com.excel.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;


/**
 * Created by li wen ya on 2020/8/4
 */
public class JavaExcel {

    public static void main(String[] args) {
        String sheetName = "检测报告";
        String[] header1 = { "产品编号","", "", "", "产品名称", "", "", "" };
        String[] header2 = { "规格型号","", "", "", "订单号", "", "", "" };
        String[] header3 = { "委托单位","", "", "", "生产单位", "", "", "" };
        String[] header4 = { "生产日期","", "", "", "送样日期", "", "", "" };
        String[] header5 = { "样品数量","", "", "", "送样者", "", "", "" };
        String[] header6 = { "温度","", "湿度", "", "检测日期", "", "", "" };
        String[] header7 = { "检\r\n验\r\n依\r\n据","", "", "", "", "", "", "" };
        String[] header8 = { "产\r\n品\r\n照\r\n片","", "", "", "", "", "", "" };
        String[] header9 = { "检\r\n验\r\n结\r\n论","", "", "", "", "", "", "" };
        String[] header10 = { "批准：","", "", "审核：", "", "", "编写：", "" };
        String[] header11 = { "地址:浙江温岭市经济开发区产学研园区科技路2号","", "", "", "", "邮编：317500", "", "" };
        String[] header12 = { "电话：0576-86199103","", "", "", "", "传真：0576-86199098", "", "" };

        String[][] headers = {header1,header2,header3,header4,header5,header6,header7,header8,header9,header10,header11,header12};// 表头
        int[] columnWidth = { 12, 10, 10, 10, 10, 10, 10, 10 };// 列数
        String fileName = "国际不锈钢器皿型式检测报告";// 文件名字
        String titleName = "浙江爱仕达电器股份有限公司检测中心 ASDJC/JJ-30-04-09  A/2\r\n检　　测　　报　　告          报告编号：QT20130108-30";// 标题名字
        String result = JavaExcel.generateExcel(sheetName, titleName, columnWidth, fileName, headers);

        if (StringUtils.isBlank(result)) {
            System.out.println("导出失败");
        } else {
            System.out.println("导出成功,结果："+result);
        }
    }


    /**
     * 设置每一列的宽度
     * @param sheet 某一页
     * @param columnWidth 所有要设置的列
     */
    private static void setColumnWidth(HSSFSheet sheet, int[] columnWidth){
        for (int i = 0; i < columnWidth.length; i++) {
            for (int j = 0; j <= i; j++) {
                if (i == j) {
                    sheet.setColumnWidth(i, columnWidth[j] * 256); // 单独设置每列的宽
                }
            }
        }
    }

    /**
     *
     * @param sheetName sheet 的名字
     * @param titleName 标题名字
     * @param columnWidth 每一列的宽度
     * @param fileName 文件名字
     * @param columnName 列的名字
     * @return 报表导出 导出成功返回地址，否则返回null
     */
    private static String generateExcel(String sheetName, String titleName,
                                int[] columnWidth, String fileName, String[][] columnName) {

        // todo 创建一个workbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();

        // todo 添加一个sheet页
        HSSFSheet sheet = wb.createSheet(sheetName);

        // todo 设置每一列的宽度
        setColumnWidth(sheet,columnWidth);

        // todo 获取第一行进行设置
        HSSFRow row0 = sheet.createRow(0);
        row0.setHeightInPoints(70);// 设置单元格的高度（标题）

        // todo 创建标题单元格样式以及字体样式
        HSSFCellStyle style = getCellStyle(wb);

        // todo 获取字体与大小样式
        HSSFFont blackBold15 = getHSSFFont(wb,true,"黑体",15);// 黑体加粗15号
        HSSFFont romanBold8 = getHSSFFont(wb,true,"Times New Roman",8);// 罗马加粗8号
        HSSFFont songTypeface10 = getHSSFFont(wb,false,"宋体",10);// 宋体10号
        HSSFFont roman8 = getHSSFFont(wb,false,"Times New Roman",8);// 罗马8号
        HSSFFont roman12 = getHSSFFont(wb,false,"Times New Roman",12);// 罗马12号
        HSSFFont songTypeface12 = getHSSFFont(wb,false,"宋体",12);// 宋体12号
        // 获取要设置文字的数量
        int head1 = titleName.trim().indexOf(" ");
        int head2 = titleName.trim().indexOf("\r\n");
        int head3 = titleName.trim().indexOf("          ");
        int head4 = titleName.trim().indexOf("：");
        // 设置文字多样式
        HSSFRichTextString ts = new HSSFRichTextString(titleName);
        ts.applyFont(0,head1,blackBold15);
        ts.applyFont(head1,head2,romanBold8);
        ts.applyFont(head2,head3,blackBold15);
        ts.applyFont(head3,head4,songTypeface10);
        ts.applyFont(head4,ts.length(),roman8);

        // todo 创建第一列
        HSSFCell cell1 = row0.createCell(0);
        cell1.setCellValue(ts);// 设置值（标题）
        cell1.setCellStyle(style);// 设置标题样式

        // todo 获取每一行并设置行高
        HSSFRow row1 = getRowHeightInPoints(sheet,1,30);
        HSSFRow row2 = getRowHeightInPoints(sheet,2,30);
        HSSFRow row3 = getRowHeightInPoints(sheet,3,30);
        HSSFRow row4 = getRowHeightInPoints(sheet,4,30);
        HSSFRow row5 = getRowHeightInPoints(sheet,5,30);
        HSSFRow row6 = getRowHeightInPoints(sheet,6,30);
        HSSFRow row7 = getRowHeightInPoints(sheet,7,60);
        HSSFRow row8 = getRowHeightInPoints(sheet,8,60);
        HSSFRow row9 = getRowHeightInPoints(sheet,9,60);
        HSSFRow row10 = getRowHeightInPoints(sheet,10,40);
        HSSFRow row11 = getRowHeightInPoints(sheet,11,30);
        HSSFRow row12 = getRowHeightInPoints(sheet,12,30);

        HSSFRow[] rows = {row1,row2,row3,row4,row5,row6,row7,row8,row9,row10,row11,row12};

        // todo 获取表头单元格样式
        style = getCellStyleBorder(wb);

        // 创建字体样式
        HSSFFont headerFont = getHSSFFont(wb,false,"黑体",10);
        style.setFont(headerFont);// 为标题样式设置字体样式

        // todo 创建表头的列
        for(int j = 0; j<columnName.length; j++){
            String[] arr = columnName[j];
            for (int i = 0; i < 8; i++) {
                if(j > 8){
                    HSSFCellStyle styleLeft = getCellStyle(wb,HorizontalAlignment.LEFT);
                    if(j>9){
                        if(j == 10){
                            if(i == 0){
                                HSSFCell cell = rows[j].createCell(i);// 创建列
                                HSSFRichTextString ts1 = new HSSFRichTextString(arr[i]);
                                ts1.applyFont(0,arr[i].length()-2,songTypeface12);
                                ts1.applyFont(arr[i].length()-2,arr[i].length()-1,roman12);
                                ts1.applyFont(arr[i].length()-1,arr[i].length(),songTypeface12);
                                cell.setCellValue(ts1);// 设置单元格内容
                                cell.setCellStyle(styleLeft);// 设置每个单元格的样式
                            } else if(i == 5){
                                setFaxPhoneEmailStyle(roman12, songTypeface12, rows[j].createCell(i), arr[i], styleLeft);
                            }
                        }
                        if(j == 11){
                            if(i == 0){
                                setFaxPhoneEmailStyle(roman12, songTypeface12, rows[j].createCell(i), arr[i], styleLeft);
                            } else if(i == 5){
                                setFaxPhoneEmailStyle(roman12, songTypeface12, rows[j].createCell(i), arr[i], styleLeft);
                            }
                        }

                    } else {
                        HSSFCell cell = rows[j].createCell(i);// 创建列
                        cell.setCellValue(arr[i]);// 设置单元格内容
                        cell.setCellStyle(styleLeft);// 设置每个单元格的样式
                    }
                } else {
                    HSSFCell cell = rows[j].createCell(i);// 创建列
                    cell.setCellValue(arr[i]);// 设置单元格内容
                    cell.setCellStyle(style);// 设置每个单元格的样式
                }
            }
        }

        // todo 设置内容单元格合并
        sheet.addMergedRegion(new CellRangeAddress(0,0,  0, 8 - 1));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 5, 7));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 5, 7));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 5, 7));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 5, 7));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 5, 7));
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 5, 7));
        sheet.addMergedRegion(new CellRangeAddress(7, 7, 1, 7));
        sheet.addMergedRegion(new CellRangeAddress(8, 8, 1, 7));
        sheet.addMergedRegion(new CellRangeAddress(9, 9, 1, 7));

        sheet.addMergedRegion(new CellRangeAddress(11, 11, 0, 4));
        sheet.addMergedRegion(new CellRangeAddress(11, 11, 5, 7));
        sheet.addMergedRegion(new CellRangeAddress(12, 12, 0, 4));
        sheet.addMergedRegion(new CellRangeAddress(12, 12, 5, 7));

        // 文件夹路径
        String fileth1 = System.getProperty("user.dir");
        // 子文件夹名称
        String fileth2 = File.separator + "file" + File.separator + new Date().getTime() + "-";
        // 获取文件夹
        File file = new File(fileth1 + fileth2 + fileName + ".xls");
        // 获取上级文件夹
        File parentFile = file.getParentFile();
        if (!parentFile.exists() && !parentFile.isDirectory()) {
            parentFile.mkdirs();
        }
        FileOutputStream fout = null;
        try {
            fout = new FileOutputStream(file);
            wb.write(fout);
        } catch (Exception e) {
            return null;
        } finally {
            try {
                fout.close();
            } catch (IOException e) {
                return null ;
            }
        }
        return  System.getProperty("user.dir") + fileth2+fileName;

    }

    /**
     * 设置邮编、电话、传真字体样式
     * @param roman12 罗马字体12号样式
     * @param songTypeface12 宋体12号样式
     * @param cell 单元格
     * @param string 要设置的字符串
     * @param styleLeft 整个单元格的样式
     */
    private static void setFaxPhoneEmailStyle(HSSFFont roman12, HSSFFont songTypeface12, HSSFCell cell, String string, HSSFCellStyle styleLeft) {
        HSSFRichTextString ts1 = new HSSFRichTextString(string);
        ts1.applyFont(0,3,songTypeface12);
        ts1.applyFont(3, string.length(),roman12);
        cell.setCellValue(ts1);// 设置单元格内容
        cell.setCellStyle(styleLeft);// 设置每个单元格的样式
    }

    /**
     * 获取设置行高的行
     * @param sheet 在页上创建行
     * @param rowNum 设置创建的某一行
     * @param high 设置创建的行高
     * @return 返回设置好行高的行
     */
    private static HSSFRow getRowHeightInPoints(HSSFSheet sheet, int rowNum, int high) {
        HSSFRow row = sheet.createRow(rowNum);
        row.setHeightInPoints(high);
        return row;
    }

    /**
     * 获取字体样式
     * @param wb workbook用于创建书本
     * @param boo 字体是否加粗
     * @param style 字体样式如：黑体、宋体等
     * @param size 字体大小
     * @return 返回设置后的字体样式
     */
    private static HSSFFont getHSSFFont(HSSFWorkbook wb, boolean boo, String style, int size) {
        HSSFFont font = wb.createFont();
        font.setBold(boo);// 字体加粗
        font.setFontName(style);// 设置字体类型
        font.setFontHeightInPoints((short) size);// 设置字体大小
        return font;
    }


    /**
     * 获取 无边框样式水平垂直居中 填充背景线条边框
     * @param wb 要创建样式的书
     * @return 返回默认样式
     */
    private static HSSFCellStyle getCellStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);// 水平对齐（居中）
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直对齐（居中）
        style.setFillForegroundColor(new HSSFColor().getIndex());// 填充前景颜色
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);// 填充前景类型
        style.setWrapText(true);// 自动换行
        return style;
    }

    /**
     * 获取 无边框样式垂直居中 填充背景无边框 水平方向自定义
     * @param wb 创建样式的书
     * @param align 水平方向对齐方式
     * @return 返回样式
     */
    private static HSSFCellStyle getCellStyle(HSSFWorkbook wb,HorizontalAlignment align) {
        HSSFCellStyle style = wb.createCellStyle();// 用书创建一个样式
        style.setAlignment(align);// 水平对齐（居中）
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直对齐（居中）
        style.setFillForegroundColor(new HSSFColor().getIndex());// 填充前景颜色
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);// 填充前景类型
        style.setWrapText(true);// 自动换行
        return style;
    }



    /**
     * 创建 有边框样式水平垂直居中 无填充背景
     * @param wb 用于创建单元格样式
     * @return 返回样式
     */
    private static HSSFCellStyle getCellStyleBorder(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);// 水平居中对齐
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直居中对齐
        style.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());// 边框颜色（黑色）
        style.setBorderBottom(BorderStyle.THIN);// 底部边框
        style.setBorderLeft(BorderStyle.THIN);// 左边框
        style.setBorderRight(BorderStyle.THIN);// 右边框
        style.setBorderTop(BorderStyle.THIN);// 顶部边框
        style.setWrapText(true);// 设置自动换行
        return style;
    }

    /**
     * 创建 有边框样式水平垂直居中 无填充背景
     * @param wb 用于创建单元格样式
     * @return 返回样式
     */
    private static HSSFCellStyle getCellStyleBorder(HSSFWorkbook wb,HorizontalAlignment align) {
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(align);// 水平居中对齐
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直居中对齐
        style.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());// 边框颜色（黑色）
        style.setBorderBottom(BorderStyle.THIN);// 底部边框
        style.setBorderLeft(BorderStyle.THIN);// 左边框
        style.setBorderRight(BorderStyle.THIN);// 右边框
        style.setBorderTop(BorderStyle.THIN);// 顶部边框
        style.setWrapText(true);// 设置自动换行
        return style;
    }


}
