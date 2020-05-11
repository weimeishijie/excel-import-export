package com.excel.input;

import cn.afterturn.easypoi.entity.vo.MapExcelConstants;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.view.PoiBaseView;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by li wen ya on 2020/5/9
 */
@Controller
@RequestMapping("/excel")
public class ExportExcelController {

    /**
     * excel导出
     *
     * @author fengshuonan
     * @Date 2019/3/9 11:03
     */
    @GetMapping
    public void export(ModelMap modelMap, HttpServletRequest request,
                       HttpServletResponse response) {
        //初始化表头
        List<ExcelExportEntity> entity = new ArrayList<>();
        entity.add(new ExcelExportEntity("用户id", "user_id"));
        entity.add(new ExcelExportEntity("头像", "avatar"));
        entity.add(new ExcelExportEntity("账号", "account"));
        entity.add(new ExcelExportEntity("姓名", "name"));
        entity.add(new ExcelExportEntity("生日", "birthday"));
        entity.add(new ExcelExportEntity("性别", "sex"));
        entity.add(new ExcelExportEntity("邮箱", "email"));
        entity.add(new ExcelExportEntity("电话", "phone"));
        entity.add(new ExcelExportEntity("角色id", "role_id"));
        entity.add(new ExcelExportEntity("部门id", "dept_id"));
        entity.add(new ExcelExportEntity("状态", "status"));
        entity.add(new ExcelExportEntity("创建时间", "create_time"));

        //初始化化数据
//        List<Map<String, Object>> maps = userService.listMaps();
        List<Map<String, Object>> maps = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("hello", "world");
        maps.add(map);
        ArrayList<Map<String, Object>> total = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            total.addAll(maps);
        }

        // title: 表的标题 sheetName: 表的sheet名
        ExportParams params = new ExportParams("Guns管理系统所有用户", "用户表", ExcelType.XSSF);
        modelMap.put(MapExcelConstants.MAP_LIST, total);
        modelMap.put(MapExcelConstants.ENTITY_LIST, entity);// 内容字段的表头
        modelMap.put(MapExcelConstants.PARAMS, params);// 参数列表 设置表头明与sheet名
        modelMap.put(MapExcelConstants.FILE_NAME, "Guns管理系统所有用户");
        PoiBaseView.render(modelMap, request, response, MapExcelConstants.EASYPOI_MAP_EXCEL_VIEW);
    }

}
