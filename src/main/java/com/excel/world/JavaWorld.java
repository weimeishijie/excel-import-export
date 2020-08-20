package com.excel.world;


import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by li wen ya on 2020/8/20
 */
public class JavaWorld {

    private static String path = System.getProperty("user.dir")+"\\file\\";

    public static void main(String[] args) throws Exception {
        Map<String, Object> data = getTemplateData();
        String fileName = "检测报告模板-10条.docx";
        Map<String, Object> picture = getTemplatePicture();
        getWord(data,picture,fileName);
    }

    private static Map<String, Object> getTemplatePicture(){
        Map<String, Object> pic = new HashMap<>();
        pic.put("${t19}", path+"\\qm.png");
        return pic;
    }

    private static Map<String, Object> getTemplateData(){
        Map<String, Object> data = new HashMap<>();
        data.put("${t1}", "hello");
        data.put("${t2}", "world");
        data.put("${t3}", "test");
        data.put("${t4}", "template");
        data.put("${t5}", "test");
        data.put("${t6}", "test");
        data.put("${t7}", "test");
        data.put("${t8}", "test");
        data.put("${t9}", "test");
        data.put("${t10}", "test");
        data.put("${t11}", "test");
        data.put("${t12}", "test");
        data.put("${t13}", "test");
        data.put("${t14}", "test");
        data.put("${t15}", "test");
        data.put("${t16}", "test");
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        String date = sdf.format(new Date());
        data.put("${t17}", date);
        data.put("${t18}", "test");
//        data.put("${t19}", "test");
        data.put("${t20}", "test");
        data.put("${t21}", "test");
        data.put("${t22}", "test");
        data.put("${r1_1}", "test");
        data.put("${r1_2}", "test");
        data.put("${r1_3}", "test");
        data.put("${r1_4}", "test");
        data.put("${r1_5}", "test");
        data.put("${r1_6}", "test");
        data.put("${r2_1}", "test");
        data.put("${r2_2}", "test");
        data.put("${r2_3}", "test");
        data.put("${r2_4}", "test");
        data.put("${r2_5}", "test");
        data.put("${r2_6}", "test");
        data.put("${r3_1}", "test");
        data.put("${r3_2}", "test");
        data.put("${r3_3}", "test");
        data.put("${r3_4}", "test");
        data.put("${r3_5}", "test");
        data.put("${r3_6}", "test");
        data.put("${r4_1}", "test");
        data.put("${r4_2}", "test");
        data.put("${r4_3}", "test");
        data.put("${r4_4}", "test");
        data.put("${r4_5}", "test");
        data.put("${r4_6}", "test");
        data.put("${r5_1}", "test");
        data.put("${r5_2}", "test");
        data.put("${r5_3}", "test");
        data.put("${r5_4}", "test");
        data.put("${r5_5}", "test");
        data.put("${r5_6}", "test");
        data.put("${r6_1}", "test");
        data.put("${r6_2}", "test");
        data.put("${r6_3}", "test");
        data.put("${r6_4}", "test");
        data.put("${r6_5}", "test");
        data.put("${r6_6}", "test");
        data.put("${r7_1}", "test");
        data.put("${r7_2}", "test");
        data.put("${r7_3}", "test");
        data.put("${r7_4}", "test");
        data.put("${r7_5}", "test");
        data.put("${r7_6}", "test");
        data.put("${r8_1}", "test");
        data.put("${r8_2}", "test");
        data.put("${r8_3}", "test");
        data.put("${r8_4}", "test");
        data.put("${r8_5}", "test");
        data.put("${r8_6}", "test");
        data.put("${r9_1}", "test");
        data.put("${r9_2}", "test");
        data.put("${r9_3}", "test");
        data.put("${r9_4}", "test");
        data.put("${r9_5}", "test");
        data.put("${r9_6}", "test");
        data.put("${r10_1}", "test");
        data.put("${r10_2}", "test");
        data.put("${r10_3}", "test");
        data.put("${r10_4}", "test");
        data.put("${r10_5}", "test");
        data.put("${r10_6}", "test");
        return data;
    }

    private static void getWord(Map<String, Object> data, String fileName) throws Exception {
        try (FileInputStream is = new FileInputStream(path+fileName);
             XWPFDocument document = new XWPFDocument(is)) {
            // 替换掉表格之外的文本(仅限文本)
            changeText(document, data);

            // 替换表格内的文本对象
            changeTableText(document, data);

            long time = System.currentTimeMillis();// 获取系统时间
            System.out.println(time); // 打印时间
            // 使用try和catch关键字捕获异常
            try (FileOutputStream out = new FileOutputStream(path+"检测报告模板"+new Date().getTime()+ ".docx")) {
                document.write(out);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    private static void getWord(Map<String, Object> data, Map<String, Object> picture, String fileName) throws Exception {
        try (FileInputStream is = new FileInputStream(path+fileName);
             XWPFDocument document = new XWPFDocument(is)) {
            // 替换掉表格之外的文本(仅限文本)
            changeText(document, data);

            // 替换表格内的文本对象
            changeTableText(document, data);

            // 替换图片
            changePic(document, picture);

            // 替换表格内的图片对象
//            changeTablePic(document, picture);

            long time = System.currentTimeMillis();// 获取系统时间
            System.out.println(time); // 打印时间
            // 使用try和catch关键字捕获异常
            try (FileOutputStream out = new FileOutputStream(path+"检测报告模板"+new Date().getTime() + ".docx")) {
                document.write(out);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    /**
     * 替换段落文本
     *
     * @param document docx解析对象
     * @param textMap  需要替换的信息集合
     */
    private static void changeText(XWPFDocument document, Map<String, Object> textMap) {
        // 获取段落集合 返回包含页眉或页脚文本的段落
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        // 增强型for循环语句，前面一个为声明语句，后一个为表达式
        for (XWPFParagraph paragraph : paragraphs) {
            // 判断此段落是否需要替换
            String text = paragraph.getText();// 检索文档中的所有文本
            if (checkText(text)) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    // 替换模板原来位置
                    Object ob = changeValue(run.toString(), textMap);
                    if (ob instanceof String) {
                        if (textMap.containsKey(run.toString())) {
                            run.setText((String) ob, 0);
                        }
                    }
                }
            }
        }
    }

    /* 检查文本中是否包含指定的字符(此处为“$”)，并返回值 */
    private static boolean checkText(String text) {
        boolean check = false;
        if (text.contains("$")) {
            check = true;
        }
        return check;
    }

    /**
     * 替换图片
     *
     * @param document docx解析对象
     * @param textMap 需要替换的信息
     * @throws Exception
     */
    private static void changePic(XWPFDocument document, Map<String, Object> textMap) throws Exception {
        // 获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            // 判断此段落是否需要替换
            String text = paragraph.getText();
            if (checkText(text)) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    // 替换模板原来位置
                    Object ob = changeValue(run.toString(), textMap);
                    if (ob instanceof String) {
                        if (textMap.containsKey(run.toString())) {
                            run.setText("", 0);
                            try (FileInputStream is = new FileInputStream((String) ob)) {
                                run.addPicture(is, XWPFDocument.PICTURE_TYPE_PNG, (String) ob, Units.toEMU(50), Units.toEMU(20));
                            }
                        }
                    }
                }
            }
        }
    }

    private static void changeTableText(XWPFDocument document, Map<String, Object> data) {
        // 获取文件的表格
        List<XWPFTable> tableList = document.getTables();

        // 循环所有需要进行替换的文本，进行替换
        for (XWPFTable table : tableList) {
            if (checkText(table.getText())) {
                List<XWPFTableRow> rows = table.getRows();
                // 遍历表格，并替换模板
                eachTable(rows, data);
            }
        }
    }

    private static void changeTablePic(XWPFDocument document, Map<String, Object> pic) throws Exception {
        List<XWPFTable> tableList = document.getTables();

        // 循环所有需要替换的文本，进行替换
        for (XWPFTable table : tableList) {
            if (checkText(table.getText())) {
                List<XWPFTableRow> rows = table.getRows();
                System.out.println("简单表格替换：" + rows);
                // 遍历表格，并替换模板
                eachTablePic(rows, pic);
            }
        }
    }

    private static void eachTablePic(List<XWPFTableRow> rows, Map<String, Object> pic) throws Exception {
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                // 判断单元格是否需要替换
                if (checkText(cell.getText())) {
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {
                            Object ob = changeValue(run.toString(), pic);
                            if (ob instanceof String) {
                                System.out.println("run" + "'" + run.toString() + "'");
                                if (pic.containsKey(run.toString())) {
                                    System.out.println("run" + run.toString() + "替换为" + ob);
                                    run.setText("", 0);
                                    try (FileInputStream is = new FileInputStream((String) ob)) {
                                        run.addPicture(is, XWPFDocument.PICTURE_TYPE_PNG, (String) ob, Units.toEMU(100),
                                                Units.toEMU(100));
                                    }
                                } else {
                                    System.out.println("'" + run.toString() + "' 不匹配");
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    private static Object changeValue(String value, Map<String, Object> textMap) {
        Set<Map.Entry<String, Object>> textSets = textMap.entrySet();
        Object valu = "";
        for (Map.Entry<String, Object> textSet : textSets) {
            // 匹配模板与替换值 格式${key}
            String key = textSet.getKey();
            if (value.contains(key)) {
                valu = textSet.getValue();
            }
        }
        return valu;
    }

    private static void eachTable(List<XWPFTableRow> rows, Map<String, Object> textMap) {
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                // 判断单元格是否需要替换
                if (checkText(cell.getText())) {
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {

                            Object ob = changeValue(run.toString(), textMap);
                            if (ob instanceof String) {
                                System.out.println("run:" + "'" + run.toString() + "'");
                                if (textMap.containsKey(run.toString())) {
                                    System.out.println("run:" + run.toString() + "替换为" + ob);
                                    run.setText((String) ob, 0);
                                } else {
                                    System.out.println("'" + run.toString() + "'不匹配");
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
