package excle;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Lenovo on 2015/6/16.
 */
public class TestExcle {

    @Test
    public void test() {
        Map map = new HashMap<String, Object>();
        map.put("name", "单位人员数据统计");
        String bigTitle[] = {"性别#-#2", "民族#-#2", "年龄#-#4", "政治面貌#-#3", "学历/学位#-#4", "技术职称#-#3", "参加工作时长#-#3"};
        map.put("bigTitle", bigTitle);
        //
        String title[] = {"部门#-#DEPART", "总人数#-#TOTAL", "男性#-#MAN", "女性#-#WOMAN", "汉族#-#HAN", "少数民族#-#SHAO", "30岁以下#-#LESS30", "30-40岁#-#THREETY2FOURTY", "40-50岁#-#FOURTY2FIVETY", "50以上#-#ABOVE50", "党员#-#DANGYUAN", "民主人士#-#MINZHU", "群众#-#QUNZHONG", "大专#-#DAZHUAN", "本科#-#DAXUE", "硕士#-#SHUOSHI", "博士#-#BOSHI", "初级职称#-#DI", "中级职称#-#ZHONG", "高级职称#-#GAO", "10年以下工作时间#-#LESSTEN", "10-20年工作时间#-#TEN2TWENTY", "20年以上工作时间#-#ABOVETWTNTY"};
        map.put("title", title);
        List<Map> list = new ArrayList<Map>();
        map.put("list", list);

        HSSFWorkbook workbook = MyExportExcle.exportForClass(map);
//        HSSFWorkbook workbook = ExportExcle.exportForClass(map);
        OutputStream outputStream = null;
        try {
            outputStream = new BufferedOutputStream(new FileOutputStream("E:\\as3.xls", false));
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.flush();
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }
}
