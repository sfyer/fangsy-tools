package com.shard.dlt;

import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.shard.dlt.bean.ExcelBean;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.BeanUtils;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Author : fsy
 * @Date: 2020-05-28 23:11
 */
public class ExcelUtil {

    private final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    private static final String METHOD = "method";
    private static final String ARGS = "args";
    private static final String SORRY_ERROR = "对不起，导入的Excel表头字段不正确，请检查模板！";

    private String errMsg = "导入文件-第%s行%s字段注入对象异常，该行导入失败";

    /**
     * excel导入产生的异常集合
     */
    private List<String> errorMsgList = new ArrayList<>();
    /**
     * 导入返回对象
     */
    private List<Object> list = new ArrayList<Object>();

    /**
     * 表头和属性的Map集合,其中Map中Key为Excel列的名称，Value为反射类的属性
     */
    private Map<String, String> titleFieldNameMap = new LinkedHashMap<>();

    /**
     * 校验规则： key为属性名称，value为校验方法以及对应的参数
     */
    private Map<String, List<Map<String, String>>> validMap = new HashMap<>();

    /**
     * 导入excel基本参数实体
     */
    private ExcelBean excelBean;


    public ExcelUtil(ExcelBean excelBean, InputStream inputStream) throws Exception {
        this.excelBean = excelBean;
        analysisData();
        readExcel(inputStream);
    }

    /**
     * @throws Exception
     * @description 将传入数据进行解析，生成规则数据
     * @author fangsy
     * @createTime 2019年7月5日下午5:30:28
     * @version 1.0.0
     */
    private void analysisData() throws Exception {
        Map<String, String> keyValueMap = this.excelBean.getKeyValueMap();
        if (keyValueMap == null || keyValueMap.isEmpty()) {
            throw new Exception("表头和属性校验方法的Map集合为空");
        }

        // 生成校验规则&&与对应属性值
        createValidRuleAndFiledName(keyValueMap);
    }

    private void createValidRuleAndFiledName(Map<String, String> keyValueMap) throws Exception {
        for (String key : keyValueMap.keySet()) {
            // 该表头对应的处理参数
            String value = keyValueMap.get(key);
            if (StringUtils.isTrimEmpty(value)) {
                throw new Exception(String.format("表头%s对应的属性名称未传值", key));
            }
            String[] valueStr = value.split(";");
            // 属性名称
            String fieldName = valueStr[0];
            titleFieldNameMap.put(key, fieldName);
            // 为属性添加校验方法
            if (valueStr.length >= 2) {
                List<Map<String, String>> rules = new ArrayList<>();
                for (int i = 1; i < valueStr.length; i++) {
                    if (StringUtils.isTrimEmpty(valueStr[i])) {
                        continue;
                    }
                    // 将一个校验方法进行解析
                    String[] methodsAndArgs = valueStr[i].split(":");

                    Map<String, String> ruleMap = new HashMap<>();
                    ruleMap.put(METHOD, methodsAndArgs[0]);
                    if (methodsAndArgs.length >= 2) {
                        if (StringUtils.isTrimEmpty(methodsAndArgs[1])) {
                            continue;
                        }
                        ruleMap.put(ARGS, methodsAndArgs[1]);
                    }
                    rules.add(ruleMap);
                }
                // 将每个属性的校验集合存放
                validMap.put(fieldName, rules);
            }
        }

    }

    /**
     * 存放每一个field字段对应所在的列的序号
     */
    private Map<String, Integer> cellmap = new HashMap<>();

    public List<String> getErrorMsgList() {
        return errorMsgList;
    }

    @SuppressWarnings("unchecked")
    public <T> List<T> getList() {
        return (List<T>) list;
    }

    /**
     * @param inputStream 输入流
     * @return List<T> 读取到的数据集合
     * @throws Exception
     * @description readExcel:根据传进来的map集合读取Excel以及model读取Excel文件
     * @author fangsy
     * @createTime 2019年7月2日下午2:37:33
     * @version 1.0.0
     * @since JDK 1.7
     */
    private void readExcel(InputStream inputStream) throws Exception {

        String classPath = excelBean.getClassPath();
        String fileName = excelBean.getFileName();
        Integer maxNum = excelBean.getMaxNum();
        Integer rowNumIndex = excelBean.getRowNumIndex();


        // 返回表头字段名和属性字段名Map集合中键的集合(Excel列的名称集合)
        Set<String> keySet = titleFieldNameMap.keySet();

        // 反射用
        Class<?> demo = null;
        Object obj = null;

        demo = Class.forName(classPath);
        // 获取文件名后缀判断文件类型
        String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length()).toLowerCase();

        // 根据文件类型及文件输入流新建工作簿对象
        Workbook wb = null;
        if ("xls".equals(fileType)) {
            wb = new HSSFWorkbook(inputStream);
        } else if ("xlsx".equals(fileType)) {
            wb = new XSSFWorkbook(inputStream);
        } else if("csv".equals(fileType)){
            readCsv(inputStream);
            return;
        }else {
            logger.error("您输入的excel格式不正确");
            throw new Exception("您输入的excel格式不正确");
        }
        int sheetNums = wb.getNumberOfSheets();
        if (excelBean.getSheetNums() != null) {
            sheetNums = excelBean.getSheetNums();
        }
        // 遍历每个Sheet表
        for (int sheetNum = 0; sheetNum < sheetNums; sheetNum++) {
            // 表头成功读取标志位。当表头成功读取后，rowNum_x值为表头实际行数
            int rowNum_x = -1;

            // 存放所有的表头字段信息
            List<String> headlist = new ArrayList<>();
            // 获取当前Sheet表
            Sheet hssfSheet = wb.getSheetAt(sheetNum);

            // 设置默认最大行数,当超出最大行数时返回异常
            if (hssfSheet != null && hssfSheet.getLastRowNum() > maxNum) {
                throw new Exception(String.format("excel文件：%s,sheet名称:%s-->数据超过%s行，请检查是否有空行,或分批导入", fileName,
                    hssfSheet.getSheetName(), maxNum));
            }
            if (hssfSheet != null ) {
                initPulicTitle(hssfSheet, excelBean.getPublicTitleMap());
                initPulicTitle(hssfSheet, excelBean.getPersonalcTitleMap());
            }

            if (hssfSheet != null) {
                // 遍历Excel中的每一行
                for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                    // 当表头成功读取标志位rowNum_x为-1时，说明还未开始读取数据。此时，如果传值指定读取其实行，就从指定行寻找，否则自动寻找。
                    if (rowNum_x == -1) {
                        // 判断指定行是否为空
                        Row hssfRow = hssfSheet.getRow(rowNumIndex);
                        if (hssfRow == null) {
                            throw new RuntimeException("指定的行为空，请检查");
                        }
                        // 设置当前行为指定行
                        rowNum = rowNumIndex - 1;
                    }

                    // 获取当前行
                    Row hssfRow = hssfSheet.getRow(rowNum);
                    // 当前行为空时，跳出本次循环进入下一行
                    if (hssfRow == null) {
                        continue;
                    }

                    // 获取表头内容
                    if (rowNum_x == -1) {
                        // 循环列Cell
                        for (int cellNum = 0; cellNum <= hssfRow.getLastCellNum(); cellNum++) {

                            Cell hssfCell = hssfRow.getCell(cellNum);
                            // 当前cell为空时，跳出本次循环，进入下一列。
                            if (hssfCell == null) {
                                continue;
                            }
                            hssfCell.setCellType(Cell.CELL_TYPE_STRING);
                            // 获取当前cell的值(String类型)
                            String tempCellValue = hssfSheet.getRow(rowNum).getCell(cellNum).getStringCellValue();
                            // 去除空格,空格ASCII码为160
                            tempCellValue = StringTool.deleteWhitespace(tempCellValue);
                            tempCellValue = tempCellValue.trim();

                            // 将表头内容放入集合
                            headlist.add(tempCellValue);
                            if (titleFieldNameMap.containsKey(tempCellValue)) {
                                cellmap.put(titleFieldNameMap.get(tempCellValue), cellNum);
                            } else {
                                if (excelBean.isValiAllTitle()) {
                                    logger.error(SORRY_ERROR);
                                    throw new Exception(SORRY_ERROR);
                                } else {
                                    continue;
                                }
                            }
                        }
                        rowNum_x = rowNumIndex;
                    } else {

                        // 实例化反射类对象
                        obj = demo.newInstance();

                        // 遍历并取出所需要的每个属性值
                        Boolean isEmptyRow = hasNext(obj, keySet, hssfRow, rowNum);
                        if(isEmptyRow && excelBean.isEndEmptyRow()){
                            break;
                        }

                    }

                }
            }

        }

        personalHandle();

        // 流关闭
        inputStream.close();
    }

    private void readCsv(InputStream inputStream) throws Exception{
        String charset = "GBK";
        try (CSVReader csvReader = new CSVReaderBuilder(new BufferedReader(new InputStreamReader(inputStream, charset))).build()) {
            Iterator<String[]> iterator = csvReader.iterator();
            int row = 0;  // 标记第几行
            int dataStartRow = this.excelBean.getRowNumIndex().intValue();
            while (iterator.hasNext()) {
                ++row;
                String[] cols = iterator.next();
                if(cols.length < 1){
                    continue;
                }
                // excel头部信息公用处理
                if(row < dataStartRow){
                    initCsvPublicData(row,cols);
                }
                // 表头初始化 对应列数和属性名
                if(row == dataStartRow){
                    initMatchCsvPojo(row,cols);
                }
                // 数据部分
                if(row > dataStartRow){
                    addPojoData(row,cols);
                }
            }

        } catch (Exception e) {

            logger.error("error{}",e);

        }
    }

    private void addPojoData(int row, String[] cols) throws Exception{
        String classPath = excelBean.getClassPath();
        Class<?> demo = Class.forName(classPath);
        Object object = demo.newInstance();
        Set<String> keySet = titleFieldNameMap.keySet();
        Object key = null;
        try {
            Iterator<String> it = keySet.iterator();
            while (it.hasNext()) {
                // Excel列名
                key = it.next();
                // 获取属性对应列数
                Integer cellNum_x = cellmap.get(titleFieldNameMap.get(key).toString());
                // 当属性对应列为空时，结束本次循环，进入下次循环，继续获取其他属性值
                String val = cols[cellNum_x];
                // 得到属性名
                String attrName = titleFieldNameMap.get(key).toString();
                // 得到属性类型
                Class<?> attrType = BeanUtils.findPropertyType(attrName, new Class[]{object.getClass()});
                val = validAll(titleFieldNameMap.get(key), val);
                setter(object, attrName, val, attrType, row, cellNum_x, key);
            }
            if (excelBean.getPublicTitleMap() != null) {
                Set<String> propertys = excelBean.getPublicTitleMap().keySet();
                for (String property : propertys) {
                    setter(object, property, excelBean.getPublicTitleMap().get(property), String.class, 0, 0, null);
                }
            }
            list.add(object);
            return ;

        } catch (ExcelAnalysisException e) {
            if (!excelBean.isSkipException()) {
                throw e;
            }
            logger.error("导入excel失败1", e);
            errorMsgList.add(String.format(e.getMessage(), row + 1, key));
        } catch (Exception e) {
            if (!excelBean.isSkipException()) {
                throw e;
            }
            logger.error("导入excel失败2", e);
            errorMsgList.add(String.format(e.getCause().getMessage(), row + 1, key));
        }
        return ;
    }

    private void initMatchCsvPojo(int row, String[] cols) throws Exception{
        for(int i = 0;i<cols.length;i++){
            String tempCellValue = cols[i].trim();
            if (titleFieldNameMap.containsKey(tempCellValue)) {
                cellmap.put(titleFieldNameMap.get(tempCellValue), i);
            } else {
                if (excelBean.isValiAllTitle()) {
                    logger.error(SORRY_ERROR);
                    throw new Exception(SORRY_ERROR);
                } else {
                    continue;
                }
            }
        }
    }

    private void initCsvPublicData(int row, String[] cols) throws Exception {
        Map<String, String> titleMap = this.excelBean.getPublicTitleMap();
        if(titleMap == null||titleMap.size() == 0){
            return;
        }

        Set<String> propertys = titleMap.keySet();
        for (String property : propertys) {
            String val = titleMap.get(property);
            if(!val.contains(",")){
                continue;
            }
            String[] arr = val.split(";");
            String value = "";
            int rowVal = Integer.parseInt(arr[0].split(",")[0]);
            if(rowVal != row){ // 如果不是想要的列 直接返回
                continue ;
            }
            int croVal = Integer.parseInt(arr[0].split(",")[1]) - 1;
            // 遍历每行内容
            for(int i = 0;i < cols.length;i++){
                if(i == croVal){
                    value = cols[i];
                }
            }
            if (arr.length > 1) {
                for (int i = 1; i < arr.length; i++) {
                    Object tmp = null;
                    if (arr[i].contains(":")) {
                        String[] arg = arr[i].split(":");
                        Method method = ExcelValidUtil.class.getMethod(arg[0], String.class, String.class);
                        tmp = method.invoke(null, value, arg[1]);
                    } else {
                        Method method = ExcelValidUtil.class.getMethod(arr[i], String.class);
                        tmp = method.invoke(null, value);
                    }
                    if (tmp != null) {
                        value = String.valueOf(tmp);
                    }
                }
            }
            titleMap.put(property, value);
        }
    }


    private void initPulicTitle(Sheet hssfSheet, Map<String, String> titleMap) throws Exception {
        if(titleMap == null){
            return;
        }
        Set<String> propertys = titleMap.keySet();
        for (String property : propertys) {
            String val = titleMap.get(property);
            String[] arr = val.split(";");
            Integer row = Integer.valueOf(arr[0].split(",")[0]) - 1;
            Integer cro = Integer.valueOf(arr[0].split(",")[1]) - 1;
            hssfSheet.getRow(row).getCell(cro).setCellType(Cell.CELL_TYPE_STRING);
            String value = hssfSheet.getRow(row).getCell(cro).getStringCellValue();
            if (arr.length > 1) {
                for (int i = 1; i < arr.length; i++) {
                    Object tmp = null;
                    if (arr[i].contains(":")) {
                        String[] arg = arr[i].split(":");
                        Method method = ExcelValidUtil.class.getMethod(arg[0], String.class, String.class);
                        tmp = method.invoke(null, value, arg[1]);
                    } else {
                        Method method = ExcelValidUtil.class.getMethod(arr[i], String.class);
                        tmp = method.invoke(null, value);
                    }
                    if (tmp != null) {
                        value = String.valueOf(tmp);
                    }
                }
            }
            titleMap.put(property, value);
        }
    }

    /**
     * @param obj     实例化对象
     * @param keySet  表头名字集合
     * @param hssfRow excel行对象
     * @param rowNum  遍历行数
     * @throws Exception
     * @description 反射为每一个excel属性赋值到实例对象
     * @author fangsy
     * @createTime 2019年7月3日下午8:16:40
     * @version 1.0.0
     */
    private Boolean hasNext(Object obj, Set<String> keySet, Row hssfRow, int rowNum) throws Exception {
        Object key = null;
        Boolean isNull = true;
        try {
            Iterator<String> it = keySet.iterator();
            while (it.hasNext()) {
                // Excel列名
                key = it.next();
                // 获取属性对应列数
                Integer cellNum_x = cellmap.get(titleFieldNameMap.get(key).toString());
                // 当属性对应列为空时，结束本次循环，进入下次循环，继续获取其他属性值
                String val = "";
                if (cellNum_x != null && hssfRow.getCell(cellNum_x) != null) {
                    // 得到属性值
                    Cell cell = hssfRow.getCell(cellNum_x);
                    // 将值转为字符串
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    val = cell.getStringCellValue();
                }
                if(StringUtils.isNotTrimEmpty(val)){
                    //说明起码有一个单元格不为null
                    isNull = false;
                }
                // 得到属性名
                String attrName = titleFieldNameMap.get(key).toString();
                // 得到属性类型
                Class<?> attrType = BeanUtils.findPropertyType(attrName, new Class[]{obj.getClass()});
                val = validAll(titleFieldNameMap.get(key), val);
                setter(obj, attrName, val, attrType, rowNum, cellNum_x, key);

            }
            if (excelBean.getPublicTitleMap() != null) {
                Set<String> propertys = excelBean.getPublicTitleMap().keySet();
                for (String property : propertys) {
                    setter(obj, property, excelBean.getPublicTitleMap().get(property), String.class, 0, 0, null);
                }
            }
            // 将实例化好并设置完属性的对象放入要返回的list中
            if (!isNull) {
                list.add(obj);
            }
        } catch (ExcelAnalysisException e) {
            if (!excelBean.isSkipException()) {
                throw e;
            }
            logger.error("导入excel失败1", e);
            if(!isNull){
                errorMsgList.add(String.format(e.getMessage(), rowNum + 1, key));
            }

        } catch (Exception e) {
            if (!excelBean.isSkipException()) {
                throw e;
            }
            logger.error("导入excel失败2", e);
            if(!isNull){
                errorMsgList.add(String.format(e.getCause().getMessage(), rowNum + 1, key));
            }
        }
        return isNull;

    }

    private void personalHandle() throws Exception {
        if (excelBean.getPersonalcTitleMap() != null) {
            Set<String> propertys = excelBean.getPersonalcTitleMap().keySet();
            for (String property : propertys) {
                setter(list.get(list.size()-1), property, excelBean.getPersonalcTitleMap().get(property), String.class, 0, 0, null);
            }
        }
    }

    /**
     * @param fieldName 属性名称
     * @param val       属性对应的某一行数据
     * @throws ExcelAnalysisException
     * @throws Exception
     * @description 对数据进行校验
     * @author fangsy
     * @createTime 2019年7月4日下午1:43:21
     * @version 1.0.0
     */
    private String validAll(String fieldName, String val) throws ExcelAnalysisException, Exception {
        try {
            if (validMap.containsKey(fieldName)) {
                List<Map<String, String>> rules = validMap.get(fieldName);
                Object temp = null;
                for (Map<String, String> ruleMap : rules) {
                    if (ruleMap.containsKey(ARGS)) {
                        Method method = ExcelValidUtil.class.getMethod(ruleMap.get(METHOD), String.class, String.class);
                        temp = method.invoke(null, val, ruleMap.get(ARGS));
                    } else {
                        Method method = ExcelValidUtil.class.getMethod(ruleMap.get(METHOD), String.class);
                        temp = method.invoke(null, val);
                    }
                    if (temp != null) {
                        val = String.valueOf(temp);
                    }
                }
            }
            return val;
        } catch (NoSuchMethodException e) {
            logger.error("被调用的方法不存在", e);
            return val;
        } catch (Exception e) {
            throw e;
        }
    }

    /**
     * @param obj       反射类对象
     * @param attrName  属性名
     * @param attrValue 属性值
     * @param attrType  属性类型
     * @param row       当前数据在Excel中的具体行数
     * @param column    当前数据在Excel中的具体列数
     * @param key       当前数据对应的Excel列名
     * @throws Exception void
     * @since JDK 1.7
     */
    public void setter(Object obj, String attrName, Object attrValue, Class<?> attrType, int row, int column, Object key)
        throws Exception {
        try {

            // 对传入值常见类型的转换
            if (attrType == String.class) {
                attrValue = String.valueOf(attrValue.toString());
            } else if (attrType == Integer.class) {
                attrValue = Double.valueOf(attrValue.toString()).intValue();
            } else if (attrType == Long.class) {
                attrValue = Double.valueOf(attrValue.toString()).longValue();
            } else if (attrType == Double.class) {
                attrValue = Double.valueOf(attrValue.toString());
            } else if (attrType == Float.class) {
                attrValue = Float.valueOf(attrValue.toString());
            } else if (attrType == Date.class) {
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                attrValue = sdf.parse(attrValue.toString());
            }
            // 获取反射的方法名
            Method method = obj.getClass().getMethod("set" + StringTool.toUpperCaseFirstOne(attrName), attrType);
            // 进行反射
            method.invoke(obj, attrValue);
        } catch (Exception e) {
            logger.error("映射对象异常:{}", e);
            throw new ExcelAnalysisException(errMsg);
        }

    }



}