package excle;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * excle到处数据
 * @author jinarn
 *
 */
public class ExportExcle {

	/**
	 * 传入实体类集合的的list导出数据
	 * @param map
	 * map的规格为
	 * name ： 文件名称，如果为空则默认用时间来命名（String）
	 * title：列名，key对应的是class的类名对应显示列名,如果为空则显示类的属性名称（String[]）
	 * List：数据内容
	 */
	public static HSSFWorkbook exportForClass(Map map){

		String flag = "success";

		String fileName = (String) map.get("name");

		if(fileName.isEmpty()){
			fileName = "export"+new Date().getTime();
		}

		String[] titleMap = (String[]) map.get("title");

		List dataList = (List) map.get("list");

		List<String> new_titleMap = new ArrayList<String>();
		List<String> new_titleMap_code = new ArrayList<String>();

		// 声明一个工作薄
		HSSFWorkbook workbook = new HSSFWorkbook();
			if(titleMap==null){//未传入列名，则由属性名称决定列名
				new_titleMap = getClassName(dataList.get(0));
				new_titleMap_code = getClassName(dataList.get(0));
			}else {
				for (String til:titleMap){
					String til2[] = til.split("#-#");
					new_titleMap.add(til2[0]);
					try {
						new_titleMap_code.add(til2[1]);
					}catch (Exception e){

					}
				}
			}
	        // 生成一个表格
	        HSSFSheet sheet = workbook.createSheet(fileName);
	        // 设置表格默认列宽度为15个字节
	        sheet.setDefaultColumnWidth((short) 15);
	        // 生成一个样式
	        HSSFCellStyle style = workbook.createCellStyle();
	        // 设置这些样式
	        style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
	        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
	        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
	        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
	        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	        // 生成一个字体
	        HSSFFont font = workbook.createFont();
	        font.setColor(HSSFColor.VIOLET.index);
	        font.setFontHeightInPoints((short) 12);
	        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
	        // 把字体应用到当前的样式
	        style.setFont(font);
	        // 生成并设置另一个样式
	        HSSFCellStyle style2 = workbook.createCellStyle();
	        style2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
	        style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	        style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
	        style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	        style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
	        style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
	        style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	        style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
	        // 生成另一个字体
	        HSSFFont font2 = workbook.createFont();
	        font2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
	        // 把字体应用到当前的样式
	        style2.setFont(font2);

	        // 声明一个画图的顶级管理器
	        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
	        // 定义注释的大小和位置,详见文档
	        HSSFComment comment = patriarch.createComment(new HSSFClientAnchor(0,
	                0, 0, 0, (short) 4, 2, (short) 6, 5));
	        // 设置注释内容
	        comment.setString(new HSSFRichTextString("可以在POI中添加注释！"));
	        // 设置注释作者，当鼠标移动到单元格上是可以在状态栏中看到该内容.
	        comment.setAuthor("leno");

	        // 产生表格标题行
	        HSSFRow row = sheet.createRow(0);
	        for (short i = 0; i < new_titleMap.size(); i++)
	        {
	            HSSFCell cell = row.createCell(i);
	            cell.setCellStyle(style);
	            HSSFRichTextString text = new HSSFRichTextString(new_titleMap.get(i));
	            cell.setCellValue(text);
	        }

	        // 遍历集合数据，产生数据行
	        int index = 0;
	       for (int j = 0; j < dataList.size(); j++) {
	            index++;
	            row = sheet.createRow(index);
	            Object t = dataList.get(j);
			   	if(t instanceof Map){
					for (int x=0;x<new_titleMap_code.size();x++){
						HSSFCell cell = row.createCell(x);
						cell.setCellStyle(style2);
						try {
							Object value = ((Map)t).get(new_titleMap_code.get(x));
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
								sheet.setColumnWidth(x, (short) (35.7 * 80));
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
							if(new_titleMap_code.get(x).toUpperCase().equals("SEX") || new_titleMap_code.get(x).toUpperCase().equals("GENDER")){
								if(value==1 || String.valueOf(value).equals("1")){
									textValue = "男";
								}else if(value==0 || String.valueOf(value).equals("0")){
									textValue = "女";
								}else{
									textValue = "未知";
								}

							}
							// 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
							if (textValue != null) {
								Pattern p = Pattern.compile("^//d+(//.//d+)?$");
								Matcher matcher = p.matcher(textValue);
								if (matcher.matches()) {
									// 是数字当作double处理
									cell.setCellValue(Double.parseDouble(textValue));
								} else {
									HSSFRichTextString richString = new HSSFRichTextString(
											textValue);
									HSSFFont font3 = workbook.createFont();
									font3.setColor(HSSFColor.BLUE.index);
									richString.applyFont(font3);
									cell.setCellValue(richString);
								}
							}
						} catch (SecurityException e) {
							e.printStackTrace();
							flag = "SecurityException";
						} catch (IllegalArgumentException e) {
							e.printStackTrace();
							flag = "IllegalArgumentException";
						} finally {
							// 清理资源
						}
					}
				}else {
					for (short i = 0; i < new_titleMap_code.size(); i++) {
						HSSFCell cell = row.createCell(i);
						cell.setCellStyle(style2);
						String getMethodName = "get"
								+ new_titleMap_code.get(i).substring(0, 1).toUpperCase()
								+ new_titleMap_code.get(i).substring(1);
						try {
							Class tCls = t.getClass();
							Method getMethod = tCls.getMethod(getMethodName,
									new Class[]
											{});
							Object value = getMethod.invoke(t, new Object[]
									{});
							// 判断值的类型后进行强制类型转换
							String textValue = null;
							// if (value instanceof Integer) {
							// int intValue = (Integer) value;
							// cell.setCellValue(intValue);
							// } else if (value instanceof Float) {
							// float fValue = (Float) value;
							// textValue = new HSSFRichTextString(
							// String.valueOf(fValue));
							// cell.setCellValue(textValue);
							// } else if (value instanceof Double) {
							// double dValue = (Double) value;
							// textValue = new HSSFRichTextString(
							// String.valueOf(dValue));
							// cell.setCellValue(textValue);
							// } else if (value instanceof Long) {
							// long longValue = (Long) value;
							// cell.setCellValue(longValue);
							// }
							if (value instanceof Boolean) {
								boolean bValue = (Boolean) value;
								textValue = "男";
								if (!bValue) {
									textValue = "女";
								}
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
							if(new_titleMap_code.get(i).toUpperCase().equals("SEX") || new_titleMap_code.get(i).toUpperCase().equals("GENDER")){
								if(value==1 || String.valueOf(value).equals("1")){
									textValue = "男";
								}else if(value==0 || String.valueOf(value).equals("0")){
									textValue = "女";
								}else{
									textValue = "未知";
								}

							}
							// 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
							if (textValue != null) {
								Pattern p = Pattern.compile("^//d+(//.//d+)?$");
								Matcher matcher = p.matcher(textValue);
								if (matcher.matches()) {
									// 是数字当作double处理
									cell.setCellValue(Double.parseDouble(textValue));
								} else {
									HSSFRichTextString richString = new HSSFRichTextString(
											textValue);
									HSSFFont font3 = workbook.createFont();
									font3.setColor(HSSFColor.BLUE.index);
									richString.applyFont(font3);
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
	private static Object getFieldValueByName(String fieldName, Object o) {
		try {
			String firstLetter = fieldName.substring(0, 1).toUpperCase();
			String getter = "get" + firstLetter + fieldName.substring(1);
			Method method = o.getClass().getMethod(getter, new Class[] {});
			Object value = method.invoke(o, new Object[] {});
			return value;
		} catch (Exception e) {
			return null;
		}
	}

	/**
	 * 获取属性名称
	 * @param obj
	 * @return
	 */
	private static List<String> getClassName(Object obj){
		Object oj = obj.getClass();
		List<String> fieldName = new ArrayList<String>();
		if(obj instanceof Map){
			Set<String> keys = ((Map) obj).keySet();
			for(String key :keys){
				fieldName.add(key);
			}
		}else{
			Field[] fields = obj.getClass().getDeclaredFields();
			for (int j = 0; j < fields.length; j++) {
				Field ff = fields[j];
				fieldName.add(ff.getName());
			}

		}

		return fieldName;
	}


}
