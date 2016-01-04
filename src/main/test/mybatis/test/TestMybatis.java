package mybatis.test;

import com.test.spider.mapper.StudentMapper;
import com.test.spider.pojo.StudentEntity;
import org.apache.ibatis.session.SqlSession;
import org.apache.ibatis.session.SqlSessionFactory;
import org.apache.ibatis.session.SqlSessionFactoryBuilder;

import java.io.InputStream;

/**
 * Created by Lenovo on 2015/6/16.
 */
public class TestMybatis {

    @org.junit.Test
    public void test() {
        //mybatis的配置文件
        String resource = "mybatis-config.xml";
        //使用类加载器加载mybatis的配置文件（它也加载关联的映射文件）
        InputStream is = TestMybatis.class.getClassLoader().getResourceAsStream(resource);
        //构建sqlSession的工厂
        SqlSessionFactory sessionFactory = new SqlSessionFactoryBuilder().build(is);
        //使用MyBatis提供的Resources类加载mybatis的配置文件（它也加载关联的映射文件）
        //Reader reader = Resources.getResourceAsReader(resource);
        //构建sqlSession的工厂
        //SqlSessionFactory sessionFactory = new SqlSessionFactoryBuilder().build(reader);
        //创建能执行映射文件中sql的sqlSession
        SqlSession session = sessionFactory.openSession();
        /**
         * 映射sql的标识字符串，
         * me.gacl.mapping.userMapper是userMapper.xml文件中mapper标签的namespace属性的值，
         * getUser是select标签的id属性值，通过select标签的id属性值就可以找到要执行的SQL
         */
//        String statement = "HttpPage.getHttpPage";//映射sql的标识字符串
        String statement = "com.test.spider.mapper.StudentMapper.getStudent";//映射sql的标识字符串
        //执行查询返回一个唯一user对象的sql
        StudentMapper studentMapper = session.getMapper(StudentMapper.class);
        StudentEntity studentEntity = null;
//        session.selectOne(statement, "12345");
        studentEntity = studentMapper.getStudent("123456");
        System.out.println(studentEntity.getStudentName());
    }
}
