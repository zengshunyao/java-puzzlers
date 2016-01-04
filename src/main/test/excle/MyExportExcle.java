package excle;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;
/**
 * Created by Lenovo on 2015/6/16.
 */

/**
 * excle到处数据
 */
public class MyExportExcle {

    /**
     * 传入实体类集合的的list导出数据
     *
     * @param map map的规格为
     *            name ： 文件名称，如果为空则默认用时间来命名（String）
     *            title：列名，key对应的是class的类名对应显示列名,如果为空则显示类的属性名称（String[]）
     *            List：数据内容
     */
    public static HSSFWorkbook exportForClass(Map map) {
        //1.校验数据并获的参数
        String flag = "success";
        String fileName = (String) map.get("name");
        if (fileName.isEmpty()) {
            fileName = "export" + new Date().getTime();
        }
        //大字段和合并单元格数量
        String[] bigtTitleMap = map.get("bigTitle") != null ? (String[]) map.get("bigTitle") : null;
        //字段名称和key映射
        String[] titleMap = (String[]) map.get("title");
        List dataList = (List) map.get("list");

        List<String> new_titleMap = new ArrayList<String>();//表头结构字段名称
        List<String> new_titleMap_code = new ArrayList<String>();//list列表map信息的key

        //2.声明一个工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();

        //3.未传入列名，则由属性名称决定列名
        if (titleMap == null) {
            new_titleMap = getClassName(dataList.get(0));
            new_titleMap_code = getClassName(dataList.get(0));
        } else {
            for (String columnKeyString : titleMap) {
                String columnKey[] = columnKeyString.split("#-#");
                new_titleMap.add(columnKey[0]);
                new_titleMap_code.add(columnKey[1]);
            }
        }
        //4.1生成一个表格
        HSSFSheet sheet = workbook.createSheet(fileName);
        //4.1.1 设置表格默认列宽度为15个字节
        sheet.setDefaultColumnWidth(15);//全局

        //4.1.2 生成一个样式作为表头
        HSSFCellStyle tableHeadCellStyle = workbook.createCellStyle();
        //4.1.3 设置这些样式的属性
        tableHeadCellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        tableHeadCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        tableHeadCellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        tableHeadCellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        tableHeadCellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        tableHeadCellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        tableHeadCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //4.1.4 生成一个字体[表头标题]
        HSSFFont tableHeadFont = workbook.createFont();
        tableHeadFont.setColor(HSSFColor.VIOLET.index);//颜色
        tableHeadFont.setFontHeightInPoints((short) 12);//高亮
        tableHeadFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//加粗
        //4.1.5把字体应用到当前的样式
        tableHeadCellStyle.setFont(tableHeadFont);

        //4.2 生成并设置另一个样式[表格内容]
        HSSFCellStyle tableBodyCellStyle = workbook.createCellStyle();
        tableBodyCellStyle.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
        tableBodyCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        tableBodyCellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        tableBodyCellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        tableBodyCellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        tableBodyCellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        tableBodyCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        tableBodyCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        //4.2.1 生成另一个字体[表格内容]
        HSSFFont tableBodyFont = workbook.createFont();
        tableBodyFont.setColor(HSSFColor.BLUE.index);//颜色
        tableBodyFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);//字体正常
        //4.2.2 把字体应用到当前[表格内容]的样式
        tableBodyCellStyle.setFont(tableBodyFont);

        //4.3 声明一个画图的顶级管理器
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        //4.4 定义单元格注释的大小和位置,详见文档
        HSSFComment comment = patriarch.createComment(new HSSFClientAnchor(0,
                0, 0, 0, (short) 4, 2, (short) 6, 5));
        //4.4.1 设置注释内容
        comment.setString(new HSSFRichTextString("可以在POI中添加注释！"));
        //4.4.2设置注释作者，当鼠标移动到单元格上是可以在状态栏中看到该内容.
        comment.setAuthor("lenovo");

        /************************************************************************************
         * 此处添加合并单元格[大字段]
         ************************************************************************************/
        int index = 0;//行号
        HSSFRow row = null;
        HSSFRow row1 = sheet.createRow(index++);
//        index++;
        int start = 0;
        if (bigtTitleMap != null) {
            HSSFRow row2 = sheet.createRow(index++);//创建第一行
//            index++;
            sheet.addMergedRegion(new CellRangeAddress(index - 2, index - 1, 0, 0));//部门
            sheet.addMergedRegion(new CellRangeAddress(index - 2, index - 1, 1, 1));//总统计

            HSSFCell cell1 = row1.createCell(0);//创建单元格
            cell1.setCellStyle(tableHeadCellStyle);//设置单元格样式
            cell1.setCellValue(new HSSFRichTextString("部门"));//创建文字标题,设置文字标题到单元格

            HSSFCell cell2 = row1.createCell(1);//创建单元格
            cell2.setCellStyle(tableHeadCellStyle);//设置单元格样式
            cell2.setCellValue(new HSSFRichTextString("总统计"));//创建文字标题,设置文字标题到单元格
            int begin = 1, end = 1;
            start = 2;
            for (int i = 0; i < bigtTitleMap.length; i++) {
                String nameNumString = bigtTitleMap[i];
                String[] nameNum = nameNumString.split("#-#");
                String name = nameNum[0];
                String num = nameNum[1];
                begin = end + 1;
                end = begin + Integer.valueOf(num) - 1;
                sheet.addMergedRegion(new CellRangeAddress(index - 2, index - 2, begin, end));
                HSSFCell cell = row1.createCell(begin);//创建单元格
                cell.setCellStyle(tableHeadCellStyle);//设置单元格样式
                HSSFRichTextString text = new HSSFRichTextString(name);//创建文字标题
                cell.setCellValue(text);//设置文字标题到单元格
            }
            row = row2;
        } else {
            row = row1;
        }
        /**********************************************************************************
         * 此处添加合并单元格[小字段]
         **********************************************************************************/
        //4.5 添加表格表头标题行
//        row = sheet.createRow(index);
//        if (start == 2) {
//            return workbook;
//        }
        for (int i = start; i < new_titleMap.size(); i++) {
            HSSFCell cell = row.createCell(i);//创建单元格
            cell.setCellStyle(tableHeadCellStyle);//设置单元格样式
            HSSFRichTextString text = new HSSFRichTextString(new_titleMap.get(i));//创建文字标题
            cell.setCellValue(text);//设置文字标题到单元格
        }

        //sheet.addMergedRegion(new CellRangeAddress(index-1, index, 0, 0));//
        //sheet.addMergedRegion(new CellRangeAddress(index-1, index, 1, 1));//
        //4.6 遍历集合数据，产生数据行
        //4.6.1 校验数据
        if (dataList == null || dataList.size() < 1) {//校验数据集合
            return workbook;
        }
        //4.6.1 遍历数据添加数据到表内容
        for (int j = 0; j < dataList.size(); j++, index++) {
            row = sheet.createRow(index);
            Object dataObject = dataList.get(j);
            //4.6.1.1 如果是map
            if (dataObject instanceof Map) {
                for (int x = 0; x < new_titleMap_code.size(); x++) {
                    HSSFCell cell = row.createCell(x);
                    cell.setCellStyle(tableBodyCellStyle);

//                    try {
                    Object value = ((Map) dataObject).get(new_titleMap_code.get(x));
                    // 判断值的类型后进行强制类型转换
                    String textValue = null;
                    if (value instanceof Boolean) {
                        boolean bValue = (Boolean) value;
                        textValue = "是";
                        if (!bValue) {
                            textValue = "否";
                        }
                    } else if (value instanceof Date) {
                        Date date = (Date) value;
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        textValue = sdf.format(date);
                    } else if (value instanceof byte[]) {
                        // 有图片时，设置行高为60px;
                        row.setHeightInPoints(60);
                        // 设置图片所在列宽度为80px,注意这里单位的一个换算
                        sheet.setColumnWidth(x, (int) (35.7 * 80));//单独设置列宽度
                        // sheet.autoSizeColumn(i);
                        byte[] bsValue = (byte[]) value;
                        HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0,
                                1023, 255, (short) 6, index, (short) 6, index);
                        anchor.setAnchorType(2);
                        patriarch.createPicture(anchor, workbook.addPicture(
                                bsValue, HSSFWorkbook.PICTURE_TYPE_JPEG));
                    } else {
                        // 其它数据类型都当作字符串简单处理
                        textValue = String.valueOf(value);
                    }
                    if (new_titleMap_code.get(x).toUpperCase().equals("SEX") || new_titleMap_code.get(x).toUpperCase().equals("GENDER")) {
                        if (value == 1 || String.valueOf(value).equals("1")) {
                            textValue = "男";
                        } else if (value == 0 || String.valueOf(value).equals("0")) {
                            textValue = "女";
                        } else {
                            textValue = "未知";
                        }
                    }
                    // 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
                    if (textValue != null) {
                        String regexp = new String("^//d+(//.//d+)?$");
                        if (textValue.matches(regexp)) {
                            // 是数字当作double处理
                            cell.setCellValue(Double.parseDouble(textValue));
                        } else {
                            HSSFRichTextString richString = new HSSFRichTextString(
                                    textValue);
                            // HSSFFont font3 = workbook.createFont();
                            // font3.setColor(HSSFColor.BLUE.index);
                            // richString.applyFont(font3);
                            cell.setCellValue(richString);
                            ////cell.setCellStyle(tableBodyCellStyle);
                        }
                    }
//                    } catch (SecurityException e) {
//                        e.printStackTrace();
//                        flag = "SecurityException";
//                    } catch (IllegalArgumentException e) {
//                        e.printStackTrace();
//                        flag = "IllegalArgumentException";
//                    } finally {
//                        // 清理资源
//                    }
                }
                //4.6.1.2 不是map 是对象
            } else {
                for (int i = 0; i < new_titleMap_code.size(); i++) {
                    HSSFCell cell = row.createCell(i);
                    cell.setCellStyle(tableBodyCellStyle);
                    String getMethodName = "get"
                            + new_titleMap_code.get(i).substring(0, 1).toUpperCase()
                            + new_titleMap_code.get(i).substring(1);//获得get方法
                    try {
                        Class dataType = dataObject.getClass();
                        Method getMethod = dataType.getMethod(getMethodName, new Class[]{});
                        ////Method getMethod = dataType.getMethod(getMethodName,dataType);
                        Object value = getMethod.invoke(dataObject, new Object[]{});
                        ////Object value = getMethod.invoke(dataObject);
                        // 判断值的类型后进行强制类型转换
                        String textValue = null;
                        if (value instanceof Boolean) {
                            textValue = (Boolean) value ? "男" : "女";
                        } else if (value instanceof Date) {
                            Date date = (Date) value;
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                            textValue = sdf.format(date);
                        } else if (value instanceof byte[]) {
                            // 有图片时，设置行高为60px;
                            row.setHeightInPoints(60);
                            // 设置图片所在列宽度为80px,注意这里单位的一个换算
                            sheet.setColumnWidth(i, (short) (35.7 * 80));
                            // sheet.autoSizeColumn(i);
                            byte[] bsValue = (byte[]) value;
                            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0,
                                    1023, 255, (short) 6, index, (short) 6, index);
                            anchor.setAnchorType(2);
                            patriarch.createPicture(anchor, workbook.addPicture(
                                    bsValue, HSSFWorkbook.PICTURE_TYPE_JPEG));
                        } else {
                            // 其它数据类型都当作字符串简单处理
                            textValue = String.valueOf(value);
                        }
                        if (new_titleMap_code.get(i).toUpperCase().equals("SEX") || new_titleMap_code.get(i).toUpperCase().equals("GENDER")) {
                            if (value == 1 || String.valueOf(value).equals("1")) {
                                textValue = "男";
                            } else if (value == 0 || String.valueOf(value).equals("0")) {
                                textValue = "女";
                            } else {
                                textValue = "未知";
                            }

                        }
                        // 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
                        if (textValue != null) {
                            //Pattern p = Pattern.compile("^//d+(//.//d+)?$");
                            //Matcher matcher = p.matcher(textValue);
                            //if (matcher.matches()) {
                            String regex = new String("^//d+(//.//d+)?$");
                            if (textValue.matches(regex)) {
                                // 是数字当作double处理
                                cell.setCellValue(Double.parseDouble(textValue));
                            } else {
                                HSSFRichTextString richString = new HSSFRichTextString(
                                        textValue);
                                //HSSFFont font3 = workbook.createFont();
                                //font3.setColor(HSSFColor.BLUE.index);
                                //richString.applyFont(font3);
                                cell.setCellValue(richString);
                            }
                        }
                    } catch (SecurityException e) {
                        e.printStackTrace();
                        flag = "SecurityException";
                    } catch (NoSuchMethodException e) {
                        e.printStackTrace();
                        flag = "NoSuchMethodException";
                    } catch (IllegalArgumentException e) {
                        e.printStackTrace();
                        flag = "IllegalArgumentException";
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                        flag = "IllegalAccessException";
                    } catch (InvocationTargetException e) {
                        e.printStackTrace();
                        flag = "InvocationTargetException";

                    } finally {
                        // 清理资源
                    }
                }
            }
        }
        return workbook;
    }


    // 根据属性获取值
    private static Object getFieldValueByName(String fieldName, Object object) {
        try {
            String firstLetter = fieldName.substring(0, 1).toUpperCase();
            String getMethodName = "get" + firstLetter + fieldName.substring(1);
            Method method = object.getClass().getMethod(getMethodName, new Class[]{});
            Object value = method.invoke(object, new Object[]{});
            return value;
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取属性名称
     *
     * @param object
     * @return
     */
    private static List<String> getClassName(Object object) {
        List<String> fieldName = new ArrayList<String>();
        if (object instanceof Map) {
            Set<String> keys = ((Map) object).keySet();
            for (String key : keys) {
                fieldName.add(key);
            }
        } else {
            Field[] fields = object.getClass().getDeclaredFields();
            for (int j = 0; j < fields.length; j++) {
                Field field = fields[j];
                fieldName.add(field.getName());
            }
        }
        return fieldName;
    }
}
