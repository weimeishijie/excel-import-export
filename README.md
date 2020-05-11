# Excel 导入导出功能

### 导入功能

##### 1.代码实现：

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

##### 2.解释说明

    导入功能的实现代码放在ImportExcel类里面，并没有提取做封装，只是将代码放在了 mian 函数里面
    
### 导出功能

##### 1.代码实现

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

##### 2.解释说明

    1.此代码放在ExportExcelController控制类里面，方法实现浏览器访问接口导出到本地
    
    2.可以作为模板导出在模板内填写内容后在做导入。导入与导出封装的是同一个对象
