package com.excel.output;

import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import com.excel.output.entity.ExcelItem;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by li wen ya on 2020/5/9
 */
public class ImportExcel {

    public static void main(String[] args) {
        ObjectMapper objectMapper = new ObjectMapper();
        String rootPath = System.getProperty("user.dir");
        String name = "hello.xlsx";
        File file = new File(rootPath + File.separator+"file"+File.separator+name);
        ImportParams params = new ImportParams();
        params.setTitleRows(1);
        params.setHeadRows(1);
        List result = ExcelImportUtil.importExcel(file, ExcelItem.class, params);
        List<ExcelItem> list = new ArrayList<>();
        for (Object obj : result) {
            list.add(objectMapper.convertValue(obj, ExcelItem.class));
        }
        System.out.println(list);
        System.out.println(list.size());
    }

}
