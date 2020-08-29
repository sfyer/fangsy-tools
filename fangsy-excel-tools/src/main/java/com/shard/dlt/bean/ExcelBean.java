package com.shard.dlt.bean;

import java.util.Map;

/**
 * @Author : fsy
 * @Date: 2020-05-28 23:07
 */
public class ExcelBean {

    /**
     * Excel文件名
     */
    private String fileName;
    /**
     * 需要映射的model的路径
     */
    private String classPath;
    /**
     * 表头所在行数(从1开始，即第一行对应行数1)
     */
    private Integer rowNumIndex;
    /**
     * 最大读取行数【默认两万行】
     */
    private Integer maxNum = 20000;
    /**
     * 表头和属性校验方法的Map集合,其中Map中Key为Excel列的名称，
     * Value为反射类的属性;校验方法名:参数 ;校验方法名:参数....
     * 工具校验类映射方法名参考com.banksteel.finance.utils.excel.ExcelValidUtil
     * 例如 keyValueMap.put("ID","id;validDate:YYYY-MM-dd;notTrimEmpty")
     * <p>
     * 注:需要有序实现传入LinkedHashMap实现类
     */
    private Map<String, String> keyValueMap;
    /**
     * 公共表头的map
     * 即每条数据的字段都会赋值表头字段
     * key-字段名称,value-行数,列数
     * 例:("title":"1,1")
     * 每条数据的title都会被excel的第一列第一行的赋值
     * 读取数据后会被真实值代替("title":"上海钢银电子商务有限公司")
     */
    private Map<String, String> publicTitleMap;

    /**
     * 公共表头的map
     * 私有化的数据的字段都会赋值表头字段
     * key-字段名称,value-行数,列数
     * 例:("title":"1,1")
     * 每条数据的title都会被excel的第一列第一行的赋值
     * 读取数据后会被真实值代替("title":"上海钢银电子商务有限公司")
     */
    private Map<String, String> personalcTitleMap;


    /**
     * 是否跳过异常，继续执行下一行
     */
    private boolean isSkipException = false;
    /**
     * 是否校验所有表头,默认校验
     */
    private boolean isValiAllTitle = true;
    /**
     * 需要读取的sheet数 默认全部读取
     */
    private Integer sheetNums;
    /**
     * 碰到空行是否结束
     */
    private boolean isEndEmptyRow = true;

    public boolean isEndEmptyRow() {
        return isEndEmptyRow;
    }

    public void setEndEmptyRow(boolean endEmptyRow) {
        isEndEmptyRow = endEmptyRow;
    }

    public Integer getSheetNums() {
        return sheetNums;
    }

    public void setSheetNums(Integer sheetNums) {
        this.sheetNums = sheetNums;
    }


    public boolean isValiAllTitle() {
        return isValiAllTitle;
    }

    public void setValiAllTitle(boolean valiAllTitle) {
        this.isValiAllTitle = valiAllTitle;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getClassPath() {
        return classPath;
    }

    public void setClassPath(String classPath) {
        this.classPath = classPath;
    }

    public Integer getRowNumIndex() {
        return rowNumIndex;
    }

    public void setRowNumIndex(Integer rowNumIndex) {
        this.rowNumIndex = rowNumIndex;
    }

    public Integer getMaxNum() {
        return maxNum;
    }

    public void setMaxNum(Integer maxNum) {
        this.maxNum = maxNum;
    }

    public Map<String, String> getKeyValueMap() {
        return keyValueMap;
    }

    public void setKeyValueMap(Map<String, String> keyValueMap) {
        this.keyValueMap = keyValueMap;
    }

    public boolean isSkipException() {
        return isSkipException;
    }

    public void setSkipException(boolean isSkipException) {
        this.isSkipException = isSkipException;
    }

    public Map<String, String> getPublicTitleMap() {
        return publicTitleMap;
    }

    public void setPublicTitleMap(Map<String, String> publicTitleMap) {
        this.publicTitleMap = publicTitleMap;
    }

    public Map<String, String> getPersonalcTitleMap() {
        return personalcTitleMap;
    }

    public void setPersonalcTitleMap(Map<String, String> personalcTitleMap) {
        this.personalcTitleMap = personalcTitleMap;
    }


}