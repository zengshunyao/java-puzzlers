package spring;

import com.test.spider.controller.HttpController;
import com.test.spider.controller.HttpPageController;
import com.test.spider.httpEnum.HttpEnum;
import com.test.spider.pojo.HttpPage;
import com.test.spider.util.DateFormatUtil;
import org.apache.commons.lang3.StringUtils;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

/**
 * Created by Lenovo on 2015/6/15.
 */

public class Test {
    private ApplicationContext applicationContext = null;
    private HttpController httpController = null;
    private HttpPageController httpPageController = null;

    @BeforeClass
    public static void enter() {
        System.out.println("进来了！");
    }

    @Before
    public void init() {
        System.out.println("正在初始化。。");
        System.out.println("初始化完毕！");
        applicationContext = new ClassPathXmlApplicationContext("applicationContext.xml");
//        BeanFactoryReference bfr = DefaultLocatorFactory.getInstance().useBeanFactory("default-context");
//        BeanFactory factory = bfr.getFactory();
//        httpController = factory.getBean("httpController", HttpController.class);
//        bfr.release();
        httpController = applicationContext.getBean("httpController", HttpController.class);
        httpPageController = applicationContext.getBean("httpPageController", HttpPageController.class);
    }

    @org.junit.Test
    public void test() {
        debug();
    }

    @After
    public void destroy() {
        System.out.println("销毁对象。。。");
        System.out.println("销毁完毕！");
    }

    @AfterClass
    public static void leave() {
        System.out.println("离开了！");
    }

    public void debug() {
        String url = "http://www.baidu.com/";
        int httpType = HttpEnum.GET;
        String content = httpController.getMessage(url, httpType);
        HttpPage httpPage = new HttpPage();
        httpPage.setDomain(url);
        httpPage.setParent(null);
        httpPage.setStatus(true);
        httpPage.setTitle(StringUtils.substringBetween(content, "<title>", "</title>"));
        httpPage.setCreateDate(DateFormatUtil.toString4CurrentTimeMillis());
        httpPageController.saveHttpPage(httpPage);
        System.out.println();
    }
}
