package mybatis.basic;

import org.apache.ibatis.builder.BaseBuilder;
import org.apache.ibatis.builder.BuilderException;
import org.apache.ibatis.datasource.DataSourceFactory;
import org.apache.ibatis.io.Resources;
import org.apache.ibatis.parsing.XNode;
import org.apache.ibatis.parsing.XPathParser;
import org.apache.ibatis.session.Configuration;

import java.util.Properties;

/**
 * Created by Lenovo on 2015/6/16.
 */
public class XMLConfigBuilder extends BaseBuilder {

    private boolean parsed;
    //xml解析器
    private XPathParser parser;
    private String environment;

    public XMLConfigBuilder(Configuration configuration) {
        super(configuration);
    }

    //上次说到这个方法是在解析mybatis配置文件中能配置的元素节点
    //今天首先要看的就是properties节点和environments节点
    //特别强调 mybatis有顺序之分必须按顺序配置xml
    // configuration (properties?, settings?, typeAliases?, typeHandlers?, objectFactory?, objectWrapperFactory?, plugins?, environments?, databaseIdProvider?, mappers?)
    private void parseConfiguration(XNode root) {
        try {
            //解析properties元素
            propertiesElement(root.evalNode("properties")); //issue #117 read properties first
//            typeAliasesElement(root.evalNode("typeAliases"));
//            pluginElement(root.evalNode("plugins"));
//            objectFactoryElement(root.evalNode("objectFactory"));
//            objectWrapperFactoryElement(root.evalNode("objectWrapperFactory"));
//            settingsElement(root.evalNode("settings"));
//            //解析environments元素
//            environmentsElement(root.evalNode("environments")); // read it after objectFactory and objectWrapperFactory issue #631
//            databaseIdProviderElement(root.evalNode("databaseIdProvider"));
//            typeHandlerElement(root.evalNode("typeHandlers"));
//            mapperElement(root.evalNode("mappers"));
        } catch (Exception e) {
            throw new BuilderException("Error parsing SQL Mapper Configuration. Cause: " + e, e);
        }
    }


    //下面就看看解析properties的具体方法
    private void propertiesElement(XNode context) throws Exception {
        if (context != null) {
            //将子节点的 name 以及value属性set进properties对象
            //这儿可以注意一下顺序，xml配置优先， 外部指定properties配置其次
            Properties defaults = context.getChildrenAsProperties();
            //获取properties节点上 resource属性的值
            String resource = context.getStringAttribute("resource");
            //获取properties节点上 url属性的值, resource和url不能同时配置
            String url = context.getStringAttribute("url");
            if (resource != null && url != null) {
                throw new BuilderException("The properties element cannot specify both a URL and a resource based property file reference.  Please specify one or the other.");
            }
            //把解析出的properties文件set进Properties对象
            if (resource != null) {
                defaults.putAll(Resources.getResourceAsProperties(resource));
            } else if (url != null) {
                defaults.putAll(Resources.getUrlAsProperties(url));
            }
            //将configuration对象中已配置的Properties属性与刚刚解析的融合
            //configuration这个对象会装载所解析mybatis配置文件的所有节点元素，以后也会频频提到这个对象
            //既然configuration对象用有一系列的get/set方法， 那是否就标志着我们可以使用java代码直接配置？
            //答案是肯定的， 不过使用配置文件进行配置，优势不言而喻
            Properties vars = configuration.getVariables();
            if (vars != null) {
                defaults.putAll(vars);
            }
            //把装有解析配置propertis对象set进解析器， 因为后面可能会用到
            parser.setVariables(defaults);
            //set进configuration对象
            configuration.setVariables(defaults);
        }
    }

    //下面再看看解析enviroments元素节点的方法
    private void environmentsElement(XNode context) throws Exception {
        if (context != null) {
            if (environment == null) {
                //解析environments节点的default属性的值
                //例如: <environments default="development">
                environment = context.getStringAttribute("default");
            }
            //递归解析environments子节点
            for (XNode child : context.getChildren()) {
                //<environment id="development">, 只有enviroment节点有id属性，那么这个属性有何作用？
                //environments 节点下可以拥有多个 environment子节点
                //类似于这样： <environments default="development"><environment id="development">...</environment><environment id="test">...</environments>
                //意思就是我们可以对应多个环境，比如开发环境，测试环境等， 由environments的default属性去选择对应的enviroment
                String id = child.getStringAttribute("id");
                //isSpecial就是根据由environments的default属性去选择对应的enviroment
//                if (isSpecifiedEnvironment(id)) {
//                    //事务， mybatis有两种：JDBC 和 MANAGED, 配置为JDBC则直接使用JDBC的事务，配置为MANAGED则是将事务托管给容器，
//                    TransactionFactory txFactory = transactionManagerElement(child.evalNode("transactionManager"));
//                    //enviroment节点下面就是dataSource节点了，解析dataSource节点（下面会贴出解析dataSource的具体方法）
//                    DataSourceFactory dsFactory = dataSourceElement(child.evalNode("dataSource"));
//                    DataSource dataSource = dsFactory.getDataSource();
//                    Environment.Builder environmentBuilder = new Environment.Builder(id)
//                            .transactionFactory(txFactory)
//                            .dataSource(dataSource);
//                    //老规矩，会将dataSource设置进configuration对象
//                    configuration.setEnvironment(environmentBuilder.build());
//                }
            }
        }
    }

    //下面看看dataSource的解析方法
    private DataSourceFactory dataSourceElement(XNode context) throws Exception {
        if (context != null) {
            //dataSource的连接池
            String type = context.getStringAttribute("type");
            //子节点 name, value属性set进一个properties对象
            Properties props = context.getChildrenAsProperties();
            //创建dataSourceFactory
            DataSourceFactory factory = (DataSourceFactory) resolveClass(type).newInstance();
            factory.setProperties(props);
            return factory;
        }
        throw new BuilderException("Environment declaration requires a DataSourceFactory.");
    }
}