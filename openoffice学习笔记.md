[toc]

---













# 概述

公司有一个需求，需要处理word文件，再将word文件转换成pdf文件响应给用户，服务器是Linux服务器，代码是跑在Linux上的，有些转pdf的技术只能跑在Windows平台上。

关于Word转PDF网上能找到的方案大概有六七种，其中5种如下：

* aspose-words
* docx4j
* openoffice
* poi
* spire.doc





## aspose-words

Aspose公司旗下的最全的一套office文档管理方案，公司设在澳大利亚。

公司差不多是专做各种文件格式处理插件的，产品系列挺多

官网：https://www.aspose.com/

不需要依赖任何组件，不依赖操作系统。

收费，不考虑

```xml
<dependency>
    <groupId>com.aspose</groupId>
    <artifactId>aspose-words</artifactId>
    <version>22.11</version>
    <classifier>jdk17</classifier>
</dependency>
```









## poi

大名鼎鼎的apache的开源组件，应用非常广泛

官网：https://poi.apache.org/

组件拆分较细，引用一些类库，但都问题不大，不依赖操作系统，经常被用于excel文档处理





## OpenOffice

Apache旗下又一开源组件，前身是1998年一家德国公司StarDivision所研发出来的一个办公室软件，称之为StarOffice。1999年8月被sun公司收购。2010年团队成员分家，分出来的一批成立了新团队做一个LibreOffice。2011年6月Oracle将其捐赠给Apache基金会

OpenOffice本身就是一套Office软件，该方案需要使用jodconverter组件配合OpenOffice完成转换，推荐使用此方案



jodconverter：https://sourceforge.net/projects/jodconverter/files/





## spire.doc

收费，用公司名和邮箱可以申请一个月的试用license，不考虑





## docx4j

澳大利亚一公司赞助的开源组件

有一个开源版，还有一个Docx4j Enterprise Edition，开源版转换效果不行，表格严重错位，与原版格式严重失真

不依赖其它组件，不依赖操作系统

官网：https://www.docx4java.org/





## 其它方案

### IText

IText是直接操作PDF的，不是docx转pdf，而且java API非常难用

```xml
<dependency>
   <groupId>com.lowagie</groupId>
   <artifactId>itext</artifactId>
   <version>4.2.2</version>
</dependency>
```





### document4j

document4j是依赖于office软件，Linux无法使用









# docx4j实现

## 使用



创建项目：

![image-20231125180307179](img/openoffice学习笔记/image-20231125180307179.png)





pom文件依赖如下：

```xml
<?xml version="1.0" encoding="UTF-8"?>
<project xmlns="http://maven.apache.org/POM/4.0.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
         xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 https://maven.apache.org/xsd/maven-4.0.0.xsd">
    <modelVersion>4.0.0</modelVersion>
    <parent>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-parent</artifactId>
        <version>2.7.1</version>
        <relativePath/> <!-- lookup parent from repository -->
    </parent>
    <groupId>mao</groupId>
    <artifactId>docx4j-word-to-pdf</artifactId>
    <version>0.0.1-SNAPSHOT</version>
    <name>docx4j-word-to-pdf</name>
    <description>docx4j-word-to-pdf</description>

    <properties>
        <java.version>17</java.version>
    </properties>

    <dependencies>
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter</artifactId>
        </dependency>

        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-test</artifactId>
            <scope>test</scope>
        </dependency>

        <dependency>
            <groupId>org.docx4j</groupId>
            <artifactId>docx4j-JAXB-Internal</artifactId>
            <version>8.3.1</version>
        </dependency>
        <dependency>
            <groupId>org.docx4j</groupId>
            <artifactId>docx4j-JAXB-ReferenceImpl</artifactId>
            <version>8.3.1</version>
        </dependency>
        <dependency>
            <groupId>org.docx4j</groupId>
            <artifactId>docx4j-export-fo</artifactId>
            <version>8.3.1</version>
        </dependency>

    </dependencies>

    <build>
        <plugins>
            <plugin>
                <groupId>org.springframework.boot</groupId>
                <artifactId>spring-boot-maven-plugin</artifactId>
            </plugin>
        </plugins>
    </build>

</project>
```





编写工具类：

```java
package mao.docx4jwordtopdf.utils;

import org.apache.commons.compress.utils.IOUtils;
import org.docx4j.Docx4J;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.util.UUID;

/**
 * Project name(项目名称)：docx4j-word-to-pdf
 * Package(包名): mao.docx4jwordtopdf.utils
 * Class(类名): DocxToPdfUtils
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/11/25
 * Time(创建时间)： 18:06
 * Version(版本): 1.0
 * Description(描述)： docx转pdf
 */

public class DocxToPdfUtils
{
    private static final Logger log = LoggerFactory.getLogger(DocxToPdfUtils.class);


    /**
     * docx转pdf
     *
     * @param docxPath docx文件路径
     * @param pdfPath  pdf文件路径
     * @throws Exception 异常
     */
    public static void convertDocxToPdf(String docxPath, String pdfPath) throws Exception
    {

        FileOutputStream fileOutputStream = null;
        try
        {
            File file = new File(docxPath);
            fileOutputStream = new FileOutputStream(new File(pdfPath));
            WordprocessingMLPackage mlPackage = WordprocessingMLPackage.load(file);
            setFontMapper(mlPackage);
            Docx4J.toPDF(mlPackage, new FileOutputStream(new File(pdfPath)));
        }
        catch (Exception e)
        {
            e.printStackTrace();
            log.error("docx文档转换为PDF失败");
        }
        finally
        {
            IOUtils.closeQuietly(fileOutputStream);
        }
    }

    /**
     * 加载字体文件（解决linux环境下无中文字体问题）
     *
     * @param mlPackage {@link WordprocessingMLPackage}
     * @throws Exception 异常
     */
    private static void setFontMapper(WordprocessingMLPackage mlPackage) throws Exception
    {
        Mapper fontMapper = new IdentityPlusMapper();
        //加载字体文件（解决linux环境下无中文字体问题）
        if (PhysicalFonts.get("SimSun") == null)
        {
            System.out.println("加载本地SimSun字体库");
            //PhysicalFonts.addPhysicalFonts("SimSun", WordUtils.class.getResource("/fonts/SIMSUN.TTC"));
        }
        fontMapper.put("隶书", PhysicalFonts.get("LiSu"));
        fontMapper.put("宋体", PhysicalFonts.get("SimSun"));
        fontMapper.put("微软雅黑", PhysicalFonts.get("Microsoft Yahei"));
        fontMapper.put("黑体", PhysicalFonts.get("SimHei"));
        fontMapper.put("楷体", PhysicalFonts.get("KaiTi"));
        fontMapper.put("新宋体", PhysicalFonts.get("NSimSun"));
        fontMapper.put("华文行楷", PhysicalFonts.get("STXingkai"));
        fontMapper.put("华文仿宋", PhysicalFonts.get("STFangsong"));
        fontMapper.put("仿宋", PhysicalFonts.get("FangSong"));
        fontMapper.put("幼圆", PhysicalFonts.get("YouYuan"));
        fontMapper.put("华文宋体", PhysicalFonts.get("STSong"));
        fontMapper.put("华文中宋", PhysicalFonts.get("STZhongsong"));
        fontMapper.put("等线", PhysicalFonts.get("SimSun"));
        fontMapper.put("等线 Light", PhysicalFonts.get("SimSun"));
        fontMapper.put("华文琥珀", PhysicalFonts.get("STHupo"));
        fontMapper.put("华文隶书", PhysicalFonts.get("STLiti"));
        fontMapper.put("华文新魏", PhysicalFonts.get("STXinwei"));
        fontMapper.put("华文彩云", PhysicalFonts.get("STCaiyun"));
        fontMapper.put("方正姚体", PhysicalFonts.get("FZYaoti"));
        fontMapper.put("方正舒体", PhysicalFonts.get("FZShuTi"));
        fontMapper.put("华文细黑", PhysicalFonts.get("STXihei"));
        fontMapper.put("宋体扩展", PhysicalFonts.get("simsun-extB"));
        fontMapper.put("仿宋_GB2312", PhysicalFonts.get("FangSong_GB2312"));
        fontMapper.put("新細明體", PhysicalFonts.get("SimSun"));
        //解决宋体（正文）和宋体（标题）的乱码问题
        PhysicalFonts.put("PMingLiU", PhysicalFonts.get("SimSun"));
        PhysicalFonts.put("新細明體", PhysicalFonts.get("SimSun"));
        //宋体&新宋体
        PhysicalFont simsunFont = PhysicalFonts.get("SimSun");
        fontMapper.put("SimSun", simsunFont);
        //设置字体
        mlPackage.setFontMapper(fontMapper);
    }
}
```





编写单元测试：

```java
package mao.docx4jwordtopdf.utils;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Project name(项目名称)：docx4j-word-to-pdf
 * Package(包名): mao.docx4jwordtopdf.utils
 * Class(测试类名): DocxToPdfUtilsTest
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/11/25
 * Time(创建时间)： 18:14
 * Version(版本): 1.0
 * Description(描述)： 测试类
 */

class DocxToPdfUtilsTest
{

    @Test
    void convertDocxToPdf() throws Exception
    {
        DocxToPdfUtils.convertDocxToPdf("./test.docx","./test.pdf");
    }

    @Test
    void convertDocxToPdf2() throws Exception
    {
        //相对复杂的docx
        DocxToPdfUtils.convertDocxToPdf("./out.docx","./out.pdf");
    }
}
```





test.docx内容如下：

![image-20231125182204431](img/openoffice学习笔记/image-20231125182204431.png)





运行

![image-20231125182216678](img/openoffice学习笔记/image-20231125182216678.png)





test.pdf输出结果如下：

![image-20231125182300479](img/openoffice学习笔记/image-20231125182300479.png)





## 结论

test.pdf目前没什么大问题，字体和格式有点问题，但是out.pdf就乱了，因为out.docx是公司的合同文档，这里就不放出结果了，而且输出太慢了，CPU占用较高，会拖垮服务器的，不考虑







# poi实现

## 使用

















# openoffice

## 概述

OpenOffice.org 是一套跨平台的办公室软件套件，能在 Windows、Linux、MacOS X (X11)、和 Solaris 等操作系统上执行。它与各个主要的办公室软件套件兼容。OpenOffice.org 是开源软件，任何人都可以免费下载、使用、及推广它。

OpenOffice.org 的主要模块有 Writer (文本文档)/Calc (电子表格)/Impress (演示文稿)/Math (公式计算)/Draw (画图)/Base (数据库)

OpenOffice.org 不仅是六大组件的组合，而且与同类产品不同的是，本套件不是独立软件模块形式创建的，从一开始，它就被设计成一个完整的办公软件包







## 安装与部署

### Windows安装

#### 第一步：下载

进入下载界面：https://sourceforge.net/projects/openofficeorg.mirror/files/

![image-20231122144536092](img/openoffice学习笔记/image-20231122144536092.png)





选择一个版本

![image-20231122144553498](img/openoffice学习笔记/image-20231122144553498.png)



选择二进制包



![image-20231122144619277](img/openoffice学习笔记/image-20231122144619277.png)





选择中文简体(zh-cn)





![image-20231122144702009](img/openoffice学习笔记/image-20231122144702009.png)



下载第一个和第二个

* [Apache_OpenOffice_4.1.14_Win_x86_install_zh-CN.exe](https://sourceforge.net/projects/openofficeorg.mirror/files/4.1.14/binaries/zh-CN/Apache_OpenOffice_4.1.14_Win_x86_install_zh-CN.exe/download)
* [Apache_OpenOffice_4.1.14_Win_x86_langpack_zh-CN.exe](https://sourceforge.net/projects/openofficeorg.mirror/files/4.1.14/binaries/zh-CN/Apache_OpenOffice_4.1.14_Win_x86_langpack_zh-CN.exe/download)





![image-20231122144758573](img/openoffice学习笔记/image-20231122144758573.png)



出现这个界面，请等待几秒，即将开始下载





#### 第二步：安装

将下载完成的`Apache_OpenOffice_4.1.14_Win_x86_install_zh-CN.exe`文件双击打开

![image-20231122161407139](img/openoffice学习笔记/image-20231122161407139.png)



选择安装位置

![image-20231122161420277](img/openoffice学习笔记/image-20231122161420277.png)

![image-20231122161442635](img/openoffice学习笔记/image-20231122161442635.png)



点击下一步

![image-20231122161457537](img/openoffice学习笔记/image-20231122161457537.png)



![image-20231122161513750](img/openoffice学习笔记/image-20231122161513750.png)



下一步

![image-20231122161537275](img/openoffice学习笔记/image-20231122161537275.png)



下一步

![image-20231122161822554](img/openoffice学习笔记/image-20231122161822554.png)



下一步

![image-20231122161836416](img/openoffice学习笔记/image-20231122161836416.png)



等待安装

![image-20231122161849066](img/openoffice学习笔记/image-20231122161849066.png)



![image-20231122162010654](img/openoffice学习笔记/image-20231122162010654.png)





#### 第三步：启动GUI程序

双击打开GUI文件，界面如下：

![image-20231122162449929](img/openoffice学习笔记/image-20231122162449929.png)



![image-20231122162508511](img/openoffice学习笔记/image-20231122162508511.png)



![image-20231122162546904](img/openoffice学习笔记/image-20231122162546904.png)





程序根目录如下：

![image-20231122163748806](img/openoffice学习笔记/image-20231122163748806.png)







#### 第四步：启动服务

进入安装目录program目录下

启动服务的命令：

```sh
./soffice -headless -accept="socket,host=127.0.0.1,port=8100;urp;" -nofirststartwizard
```

或者：

```sh
./soffice -headless -accept="socket,host=0.0.0.0,port=8100;urp;" -nofirststartwizard 
```



* **host**：绑定的ip，配置为0.0.0.0，则远程ip能使用，配置为127.0.0.1，则只能本地访问
* **port**：服务启动的端口号



```sh
PS C:\Program Files (x86)\OpenOffice 4> ls


    目录: C:\Program Files (x86)\OpenOffice 4


Mode                 LastWriteTime         Length Name
----                 -------------         ------ ----
d-----        2023/11/22     16:18                help
d-----        2023/11/22     16:18                presets
d-----        2023/11/22     16:18                program
d-----        2023/11/22     16:18                readmes
d-----        2023/11/22     16:18                share
-a----          2023/2/9     18:39          10440 readme.html
-a----          2023/2/9     18:39          10017 readme.txt


PS C:\Program Files (x86)\OpenOffice 4> cd .\program\
PS C:\Program Files (x86)\OpenOffice 4\program> ./soffice -headless -accept="socket,host=0.0.0.0,port=8100;urp;" -nofirststartwizard
PS C:\Program Files (x86)\OpenOffice 4\program>
```



![image-20231122164316150](img/openoffice学习笔记/image-20231122164316150.png)



不报错，为启动成功



#### 第五步：检查是否启动成功

```sh
netstat -ano
```



![image-20231122164826687](img/openoffice学习笔记/image-20231122164826687.png)



可以看到8100端口已经在使用了，pid为46984



![image-20231122165005339](img/openoffice学习笔记/image-20231122165005339.png)



打开任务管理器，可以看到46984的确为openoffice





#### 第六步：安装中文语言包

如果下载的openoffice本来就是中文的，则没有必要再安装，可以跳过此步骤



![image-20231122172733550](img/openoffice学习笔记/image-20231122172733550.png)



![image-20231122172745153](img/openoffice学习笔记/image-20231122172745153.png)



![image-20231122172803011](img/openoffice学习笔记/image-20231122172803011.png)



![image-20231122172851825](img/openoffice学习笔记/image-20231122172851825.png)









### Linux安装











### Docker安装

















# SpringBoot整合openoffice

## 概述

因业务需要，需要使用SpringBoot将docx文件转换成PDF文件，再将PDF文件响应给用户



## docx转pdf

### 第一步：创建Springboot程序

![image-20231122165556510](img/openoffice学习笔记/image-20231122165556510.png)



![image-20231122165610503](img/openoffice学习笔记/image-20231122165610503.png)





### 第二步：配置maven

```xml
<?xml version="1.0" encoding="UTF-8"?>
<project xmlns="http://maven.apache.org/POM/4.0.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
         xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 https://maven.apache.org/xsd/maven-4.0.0.xsd">
    <modelVersion>4.0.0</modelVersion>
    <parent>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-parent</artifactId>
        <version>2.7.1</version>
        <relativePath/> <!-- lookup parent from repository -->
    </parent>
    <groupId>mao</groupId>
    <artifactId>openoffice-word-to-pdf</artifactId>
    <version>0.0.1-SNAPSHOT</version>
    <name>openoffice-word-to-pdf</name>
    <description>openoffice-word-to-pdf</description>
    <properties>
        <java.version>1.8</java.version>
    </properties>
    <dependencies>
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-web</artifactId>
        </dependency>

        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-test</artifactId>
            <scope>test</scope>
        </dependency>

        <!--spring-boot lombok-->
        <dependency>
            <groupId>org.projectlombok</groupId>
            <artifactId>lombok</artifactId>
            <optional>true</optional>
        </dependency>

        <!--spring boot 使用openoffice服务 maven依赖-->
        <!--jodconverter 核心包 -->
        <!-- https://mvnrepository.com/artifact/org.jodconverter/jodconverter-core -->
        <dependency>
            <groupId>org.jodconverter</groupId>
            <artifactId>jodconverter-core</artifactId>
            <version>4.4.6</version>
        </dependency>

        <!--springboot支持包，里面包括了自动配置类 -->
        <!-- https://mvnrepository.com/artifact/org.jodconverter/jodconverter-spring-boot-starter -->
        <dependency>
            <groupId>org.jodconverter</groupId>
            <artifactId>jodconverter-spring-boot-starter</artifactId>
            <version>4.4.6</version>
        </dependency>

        <!--jodconverter 本地支持包 -->
        <!-- https://mvnrepository.com/artifact/org.jodconverter/jodconverter-local -->
        <dependency>
            <groupId>org.jodconverter</groupId>
            <artifactId>jodconverter-local</artifactId>
            <version>4.4.6</version>
        </dependency>
    </dependencies>

    <build>
        <plugins>
            <plugin>
                <groupId>org.springframework.boot</groupId>
                <artifactId>spring-boot-maven-plugin</artifactId>
                <configuration>
                    <image>
                        <builder>paketobuildpacks/builder-jammy-base:latest</builder>
                    </image>
                </configuration>
            </plugin>
        </plugins>
    </build>

</project>

```







### 第三步：编写application.yml配置

```yaml
jodconverter:
  local:
    enabled: true
    # openoffice安装路径
    office-home: 'C:\Program Files (x86)\OpenOffice 4'
    max-tasks-per-process: 10
    #openoffice端口号
    port-numbers: 8100
```





### 第四步：启动服务

如果报以下错误：

```sh
"org.jodconverter.office.OfficeException: A process with acceptString 'socket,host=127.0.0.1,port=8100,tcpNoDelay=1;urp;StarOffice.ServiceManager' started but its pid could not be found"
```



原因：

1. 项目启动目录中带有中文路径
2. soffice的端口异常了不能正常关闭
3. 操作系统环境问题
4. 依赖版本问题





启动成功的控制台日志如下：

```sh
    .   ____          _            __ _ _
   /\\ / ___'_ __ _ _(_)_ __  __ _ \ \ \ \
  ( ( )\___ | '_ | '_| | '_ \/ _` | \ \ \ \
   \\/  ___)| |_)| | | | | || (_| |  ) ) ) )
    '  |____| .__|_| |_|_| |_\__, | / / / /
   =========|_|==============|___/=/_/_/_/
   :: Spring Boot ::                (v2.7.1)

            _____                    _____                   _______
           /\    \                  /\    \                 /::\    \
          /::\____\                /::\    \               /::::\    \
         /::::|   |               /::::\    \             /::::::\    \
        /:::::|   |              /::::::\    \           /::::::::\    \
       /::::::|   |             /:::/\:::\    \         /:::/~~\:::\    \
      /:::/|::|   |            /:::/__\:::\    \       /:::/    \:::\    \
     /:::/ |::|   |           /::::\   \:::\    \     /:::/    / \:::\    \
    /:::/  |::|___|______    /::::::\   \:::\    \   /:::/____/   \:::\____\
   /:::/   |::::::::\    \  /:::/\:::\   \:::\    \ |:::|    |     |:::|    |
  /:::/    |:::::::::\____\/:::/  \:::\   \:::\____\|:::|____|     |:::|    |
  \::/    / ~~~~~/:::/    /\::/    \:::\  /:::/    / \:::\    \   /:::/    /
   \/____/      /:::/    /  \/____/ \:::\/:::/    /   \:::\    \ /:::/    /
               /:::/    /            \::::::/    /     \:::\    /:::/    /
              /:::/    /              \::::/    /       \:::\__/:::/    /
             /:::/    /               /:::/    /         \::::::::/    /
            /:::/    /               /:::/    /           \::::::/    /
           /:::/    /               /:::/    /             \::::/    /
          /:::/    /               /:::/    /               \::/____/
          \::/    /                \::/    /                 ~~
           \/____/                  \/____/
   :: Github (https://github.com/maomao124) ::

2023-11-22 17:48:13.321  INFO 25232 --- [           main] m.o.OpenofficeWordToPdfApplication       : Starting OpenofficeWordToPdfApplication using Java 1.8.0_332 on mao with PID 25232 (D:\程序\2023Q4\openoffice-word-to-pdf\target\classes started by mao in D:\程序\2023Q4\openoffice-word-to-pdf)
2023-11-22 17:48:13.322  INFO 25232 --- [           main] m.o.OpenofficeWordToPdfApplication       : No active profile set, falling back to 1 default profile: "default"
2023-11-22 17:48:13.759  INFO 25232 --- [           main] o.s.b.w.embedded.tomcat.TomcatWebServer  : Tomcat initialized with port(s): 8080 (http)
2023-11-22 17:48:13.764  INFO 25232 --- [           main] o.apache.catalina.core.StandardService   : Starting service [Tomcat]
2023-11-22 17:48:13.764  INFO 25232 --- [           main] org.apache.catalina.core.StandardEngine  : Starting Servlet engine: [Apache Tomcat/9.0.64]
2023-11-22 17:48:13.808  INFO 25232 --- [           main] o.a.c.c.C.[Tomcat].[localhost].[/]       : Initializing Spring embedded WebApplicationContext
2023-11-22 17:48:13.808  INFO 25232 --- [           main] w.s.c.ServletWebServerApplicationContext : Root WebApplicationContext: initialization completed in 465 ms
2023-11-22 17:48:14.057  INFO 25232 --- [er-offprocmng-0] o.j.local.office.OfficeDescriptor        : soffice info (from exec path): Product: OpenOffice - Version: ??? - useLongOptionNameGnuStyle: false
2023-11-22 17:48:14.091  INFO 25232 --- [           main] o.s.b.w.embedded.tomcat.TomcatWebServer  : Tomcat started on port(s): 8080 (http) with context path ''
2023-11-22 17:48:14.095  INFO 25232 --- [           main] m.o.OpenofficeWordToPdfApplication       : Started OpenofficeWordToPdfApplication in 1.005 seconds (JVM running for 1.461)
2023-11-22 17:48:14.157  WARN 25232 --- [er-offprocmng-0] o.j.l.office.LocalOfficeProcessManager   : Profile dir 'C:\Users\mao\AppData\Local\Temp\.jodconverter_socket_host-127.0.0.1_port-8100_tcpNoDelay-1' already exists; deleting
2023-11-22 17:48:14.170  INFO 25232 --- [er-offprocmng-0] o.j.l.office.LocalOfficeProcessManager   : Starting process with --accept 'socket,host=127.0.0.1,port=8100,tcpNoDelay=1;urp;StarOffice.ServiceManager' and profileDir 'C:\Users\mao\AppData\Local\Temp\.jodconverter_socket_host-127.0.0.1_port-8100_tcpNoDelay-1'
2023-11-22 17:48:14.809  INFO 25232 --- [er-offprocmng-0] o.j.local.office.OfficeConnection        : Connected: 'socket,host=127.0.0.1,port=8100,tcpNoDelay=1'
2023-11-22 17:48:14.810  INFO 25232 --- [er-offprocmng-0] o.j.l.office.LocalOfficeProcessManager   : Started process; pid: 11904
```





### 第五步：编写单元测试



```java
package mao.openofficewordtopdf;

import org.jodconverter.core.DocumentConverter;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;

@SpringBootTest
class OpenofficeWordToPdfApplicationTests
{
    @Autowired
    private DocumentConverter documentConverter;

    @Test
    void contextLoads()
    {
        documentConverter.convert(new File("./test.docx")).to(new File("./test.pdf"));
    }

}
```



test.docx内容如下：

![image-20231122175719673](img/openoffice学习笔记/image-20231122175719673.png)





运行单元测试，日志输出如下：

```sh
2023-11-22 21:55:57.635  INFO 13336 --- [er-offprocmng-0] o.j.l.office.LocalOfficeProcessManager   : Starting process with --accept 'socket,host=127.0.0.1,port=8100,tcpNoDelay=1;urp;StarOffice.ServiceManager' and profileDir 'C:\Users\mao\AppData\Local\Temp\.jodconverter_socket_host-127.0.0.1_port-8100_tcpNoDelay-1'
2023-11-22 21:55:58.281  INFO 13336 --- [er-offprocmng-0] o.j.local.office.OfficeConnection        : Connected: 'socket,host=127.0.0.1,port=8100,tcpNoDelay=1'
2023-11-22 21:55:58.281  INFO 13336 --- [er-offprocmng-0] o.j.l.office.LocalOfficeProcessManager   : Started process; pid: 35016
2023-11-22 21:55:58.282  INFO 13336 --- [ter-poolentry-1] o.j.local.task.LocalConversionTask       : Executing local conversion task [docx -> pdf]...
```





输出的PDF文件如下：

![image-20231122215730504](img/openoffice学习笔记/image-20231122215730504.png)











文件流转换示例：

```java
DocumentFormat documentFormatSrc = documentConverter.getFormatRegistry().getFormatByExtension("docx");
        DocumentFormat documentFormatTarget = documentConverter.getFormatRegistry().getFormatByExtension("pdf");
        documentConverter.convert(
                new FileInputStream(new File("./test.docx")))
                .as(documentFormatSrc)
                .to(new FileOutputStream("./test.pdf"))
                .as(documentFormatTarget).execute();
```







### 第六步：编写文档转换服务

```java
package mao.openofficewordtopdf.service;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Project name(项目名称)：openoffice-word-to-pdf
 * Package(包名): mao.openofficewordtopdf.service
 * Interface(接口名): DocumentConverterService
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/11/23
 * Time(创建时间)： 8:49
 * Version(版本): 1.0
 * Description(描述)： 文档转换服务
 */

public interface DocumentConverterService
{

    /**
     * 文档转换
     *
     * @param sourcePath 来源路径字符串
     * @param targetPath 目标路径字符串
     */
    void converter(String sourcePath, String targetPath);

    /**
     * 文档转换
     *
     * @param sourcePath 来源路径
     * @param targetPath 目标路径
     */
    void converter(File sourcePath, File targetPath);

    /**
     * 文档转换
     *
     * @param inputStream      输入流
     * @param outputStream     输出流
     * @param sourceFileSuffix 源文件后缀
     * @param targetFileSuffix 目标文件后缀
     */
    void converter(InputStream inputStream, OutputStream outputStream,
                   String sourceFileSuffix, String targetFileSuffix);
}
```





### 第七步：编写服务实现类

```java
package mao.openofficewordtopdf.service.impl;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import mao.openofficewordtopdf.service.DocumentConverterService;
import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.document.DocumentFormat;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Project name(项目名称)：openoffice-word-to-pdf
 * Package(包名): mao.openofficewordtopdf.service.impl
 * Class(类名): DocumentConverterServiceImpl
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/11/23
 * Time(创建时间)： 8:50
 * Version(版本): 1.0
 * Description(描述)： 文档转换服务实现
 */

@Slf4j
@Service
public class DocumentConverterServiceImpl implements DocumentConverterService
{

    @Autowired
    private DocumentConverter documentConverter;

    @SneakyThrows
    @Override
    public void converter(String sourcePath, String targetPath)
    {
        documentConverter.convert(new File(sourcePath)).to(new File(targetPath)).execute();
    }

    @SneakyThrows
    @Override
    public void converter(File sourcePath, File targetPath)
    {
        documentConverter.convert(sourcePath).to(targetPath).execute();
    }

    @SneakyThrows
    @Override
    public void converter(InputStream inputStream, OutputStream outputStream, String sourceFileSuffix, String targetFileSuffix)
    {
        DocumentFormat sourceDocumentFormat = documentConverter.getFormatRegistry()
                .getFormatByExtension(sourceFileSuffix);
        DocumentFormat targetDocumentFormat = documentConverter.getFormatRegistry()
                .getFormatByExtension(targetFileSuffix);
        documentConverter
                .convert(inputStream)
                .as(sourceDocumentFormat)
                .to(outputStream)
                .as(targetDocumentFormat)
                .execute();
    }
}
```





该实现类的单元测试如下：

```java
package mao.openofficewordtopdf.service.impl;

import lombok.SneakyThrows;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Project name(项目名称)：openoffice-word-to-pdf
 * Package(包名): mao.openofficewordtopdf.service.impl
 * Class(测试类名): DocumentConverterServiceImplTest
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/11/23
 * Time(创建时间)： 9:03
 * Version(版本): 1.0
 * Description(描述)： 测试类
 */

@SpringBootTest
class DocumentConverterServiceImplTest
{
    @Autowired
    private DocumentConverterServiceImpl documentConverterService;

    @Test
    void converter()
    {
        documentConverterService.converter("./test.docx", "./test4.html");
    }

    @Test
    void testConverter()
    {
        documentConverterService.converter(new File("./test.docx"), new File("./test5.pdf"));
    }

    @SneakyThrows
    @Test
    void testConverter1()
    {
        documentConverterService.converter(new FileInputStream("./test.ppt")
                , new FileOutputStream("./test6.pdf")
                , "ppt", "pdf");
    }
}
```



经测试，一切正常





### 第八步：编写接口

```java
package mao.openofficewordtopdf.controller;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import mao.openofficewordtopdf.service.DocumentConverterService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

/**
 * Project name(项目名称)：openoffice-word-to-pdf
 * Package(包名): mao.openofficewordtopdf.controller
 * Class(类名): DocumentConverterController
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/11/23
 * Time(创建时间)： 9:13
 * Version(版本): 1.0
 * Description(描述)： 无
 */

@Slf4j
@Controller
@RequestMapping("/api")
public class DocumentConverterController
{

    @Autowired
    private DocumentConverterService documentConverterService;

    /**
     * 文档转换接口
     *
     * @param httpServletResponse HttpServletResponse对象
     * @param sourceFileSuffix    源文件后缀
     * @param targetFileSuffix    目标文件后缀
     */
    @SneakyThrows
    @PostMapping("/converter/{sourceFileSuffix}/{targetFileSuffix}")
    public void converter(HttpServletResponse httpServletResponse,
                          @PathVariable String sourceFileSuffix,
                          @PathVariable String targetFileSuffix,
                          MultipartFile file)
    {
        String originalFilename = file.getOriginalFilename();
        log.info(originalFilename);
        String[] split = originalFilename.split("\\.");
        //类型有问题
        httpServletResponse.setContentType("application/" + targetFileSuffix);
        httpServletResponse.setHeader("Content-disposition",
                "attachment;filename=" + new String((split[0] + "." + targetFileSuffix)
                        .getBytes("utf-8"), "iso-8859-1"));
        documentConverterService.converter(file.getInputStream(),
                httpServletResponse.getOutputStream(),
                sourceFileSuffix,
                targetFileSuffix);
    }
}

```





### 第九步：编写简单前端页面



```html
<!DOCTYPE html>

<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>文档转换服务</title>
</head>
<body>
<form id="from" enctype="multipart/form-data" onsubmit="return check()" method="post" action="/api/converter/docx/pdf">
    <input type="file" name="file">
    <input type="submit" name="提交">
</form>

<br>
<input id="p1" type="text" placeholder="输入的文件后缀名">
<br>
<input id="p2" type="text" placeholder="输出的文件后缀名">

<script>

    function check()
    {
        var from = document.getElementById("from")
        var p1 = document.getElementById("p1").value
        var p2 = document.getElementById("p2").value
        from.action = '/api/converter/' + p1 + '/' + p2;
        return true;
    }
</script>
</body>
</html>
```







### 第十步：启动项目测试

```sh
2023-11-24 23:25:43.069  INFO 38168 --- [           main] m.o.OpenofficeWordToPdfApplication       : Starting OpenofficeWordToPdfApplication using Java 1.8.0_332 on mao with PID 38168 (D:\程序\2023Q4\openoffice-word-to-pdf\target\classes started by mao in D:\程序\2023Q4\openoffice-word-to-pdf)
2023-11-24 23:25:43.071  INFO 38168 --- [           main] m.o.OpenofficeWordToPdfApplication       : No active profile set, falling back to 1 default profile: "default"
2023-11-24 23:25:43.441  INFO 38168 --- [           main] o.s.b.w.embedded.tomcat.TomcatWebServer  : Tomcat initialized with port(s): 8090 (http)
2023-11-24 23:25:43.444  INFO 38168 --- [           main] o.apache.catalina.core.StandardService   : Starting service [Tomcat]
2023-11-24 23:25:43.445  INFO 38168 --- [           main] org.apache.catalina.core.StandardEngine  : Starting Servlet engine: [Apache Tomcat/9.0.64]
2023-11-24 23:25:43.481  INFO 38168 --- [           main] o.a.c.c.C.[Tomcat].[localhost].[/]       : Initializing Spring embedded WebApplicationContext
2023-11-24 23:25:43.481  INFO 38168 --- [           main] w.s.c.ServletWebServerApplicationContext : Root WebApplicationContext: initialization completed in 391 ms
2023-11-24 23:25:43.639  INFO 38168 --- [er-offprocmng-0] o.j.local.office.OfficeDescriptor        : soffice info (from exec path): Product: OpenOffice - Version: ??? - useLongOptionNameGnuStyle: false
2023-11-24 23:25:43.725  INFO 38168 --- [           main] o.s.b.a.w.s.WelcomePageHandlerMapping    : Adding welcome page: class path resource [static/index.html]
2023-11-24 23:25:43.777  INFO 38168 --- [er-offprocmng-0] o.j.l.office.LocalOfficeProcessManager   : Starting process with --accept 'socket,host=127.0.0.1,port=8100,tcpNoDelay=1;urp;StarOffice.ServiceManager' and profileDir 'C:\Users\mao\AppData\Local\Temp\.jodconverter_socket_host-127.0.0.1_port-8100_tcpNoDelay-1'
2023-11-24 23:25:43.785  INFO 38168 --- [           main] o.s.b.w.embedded.tomcat.TomcatWebServer  : Tomcat started on port(s): 8090 (http) with context path ''
2023-11-24 23:25:43.789  INFO 38168 --- [           main] m.o.OpenofficeWordToPdfApplication       : Started OpenofficeWordToPdfApplication in 0.889 seconds (JVM running for 1.419)
2023-11-24 23:25:44.456  INFO 38168 --- [er-offprocmng-0] o.j.local.office.OfficeConnection        : Connected: 'socket,host=127.0.0.1,port=8100,tcpNoDelay=1'
2023-11-24 23:25:44.457  INFO 38168 --- [er-offprocmng-0] o.j.l.office.LocalOfficeProcessManager   : Started process; pid: 10832
```



访问：http://localhost:8090/



选择文件名测试

![image-20231124232659774](img/openoffice学习笔记/image-20231124232659774.png)



![image-20231124232714865](img/openoffice学习笔记/image-20231124232714865.png)





转换无问题

![image-20231124232756001](img/openoffice学习笔记/image-20231124232756001.png)



























# java调用openoffice

## 概述

并不是所有的项目都是springboot项目，下面将不使用springboot项目来调用openoffice服务







## docx转pdf

### 第一步：新建maven项目

![image-20231123104358567](img/openoffice学习笔记/image-20231123104358567.png)





这里项目名称为`openoffice-word-to-pdf2`





### 第二步：配置maven

依赖如下：

```xml
<dependencies>
        <dependency>
            <groupId>org.openoffice</groupId>
            <artifactId>juh</artifactId>
            <version>4.1.2</version>
        </dependency>
        <dependency>
            <groupId>org.openoffice</groupId>
            <artifactId>jurt</artifactId>
            <version>4.1.2</version>
        </dependency>
        <dependency>
            <groupId>org.openoffice</groupId>
            <artifactId>ridl</artifactId>
            <version>4.1.2</version>
        </dependency>
        <dependency>
            <groupId>org.openoffice</groupId>
            <artifactId>unoil</artifactId>
            <version>4.1.2</version>
        </dependency>
        <dependency>
            <groupId>com.artofsolving</groupId>
            <artifactId>jodconverter</artifactId>
            <version>2.2.2</version>
        </dependency>
    </dependencies>
```



jodconverter需要单独下载，因为中央仓库里没有

![image-20231123104545185](img/openoffice学习笔记/image-20231123104545185.png)





下载地址：https://sourceforge.net/projects/jodconverter/files/



![image-20231123104624984](img/openoffice学习笔记/image-20231123104624984.png)



![image-20231123104633790](img/openoffice学习笔记/image-20231123104633790.png)



点击2.2.2版本

![image-20231123104654141](img/openoffice学习笔记/image-20231123104654141.png)



下载：[jodconverter-2.2.2.zip](https://sourceforge.net/projects/jodconverter/files/JODConverter/2.2.2/jodconverter-2.2.2.zip/download)



下载完成后，解压，可以得到这样的目录：

![image-20231123104806112](img/openoffice学习笔记/image-20231123104806112.png)





在lib目录下运行：

```sh
mvn install:install-file -Dfile="jodconverter-2.2.2.jar" -DgroupId=com.artofsolving -DartifactId=jodconverter -Dversion=2.2.2 -Dpackaging=jar
```





或者将lib目录拷贝到项目目录下

![image-20231123105254900](img/openoffice学习笔记/image-20231123105254900.png)



然后再打开项目结构

![image-20231123105318118](img/openoffice学习笔记/image-20231123105318118.png)



选择库

![image-20231123105407924](img/openoffice学习笔记/image-20231123105407924.png)



点击新建，选择lib包目录

![image-20231123105444088](img/openoffice学习笔记/image-20231123105444088.png)

![image-20231123105456354](img/openoffice学习笔记/image-20231123105456354.png)







### 第三步：编写DocumentConverterUtil

```java
package mao;

import com.artofsolving.jodconverter.DefaultDocumentFormatRegistry;
import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.DocumentFormat;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.StreamOpenOfficeDocumentConverter;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.ConnectException;
import java.util.Properties;

/**
 * Project name(项目名称)：openoffice-word-to-pdf2
 * Package(包名): mao
 * Class(类名): DocumentConverterUtil
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/11/24
 * Time(创建时间)： 23:53
 * Version(版本): 1.0
 * Description(描述)： 文档转换工具
 */

public class DocumentConverterUtil
{
    private static SocketOpenOfficeConnection connection = null;

    private static DocumentConverter documentConverter = null;

    private static Properties properties = null;

    private static int serviceCheckCycle;

    static
    {
        InputStream inputStream = DocumentConverterUtil.class.getClassLoader().getResourceAsStream("conf/openoffice.properties");
        try
        {
            properties = new Properties();
            properties.load(inputStream);
            String host = properties.getProperty("server.host");
            Integer port = Integer.parseInt(properties.getProperty("server.port"));
            serviceCheckCycle = Integer.parseInt(properties.getProperty("serviceCheckCycle"));
            connection = new SocketOpenOfficeConnection(host, port);
            connection.connect();
            documentConverter = new StreamOpenOfficeDocumentConverter(connection);
        }
        catch (ConnectException e)
        {
            e.printStackTrace();
            System.out.println("文档转换服务连接失败");
            //throw new RuntimeException(e);
        }
        catch (IOException e)
        {
            e.printStackTrace();
            System.out.println("加载配置文件失败或者其它问题");
            //throw new RuntimeException(e);
        }
        finally
        {
            //服务检查
            new Thread(new Runnable()
            {
                @Override
                public void run()
                {
                    while (true)
                    {
                        try
                        {
                            Thread.sleep(serviceCheckCycle);
                            if (isNotConnection())
                            {
                                System.out.println("openoffice服务异常，正在尝试重连");
                                retry();
                            }
                        }
                        catch (InterruptedException e)
                        {
                            e.printStackTrace();
                        }
                    }
                }
            }).start();

            Runtime.getRuntime().addShutdownHook(new Thread(() ->
            {
                if (connection.isConnected())
                {
                    System.out.println("正在关闭openoffice连接...");
                    connection.disconnect();
                }
            }));
        }
    }

    /**
     * 服务重连
     *
     * @return 重连结果
     */
    public static boolean retry()
    {
        System.out.println("尝试重连openoffice");
        String host = properties.getProperty("server.host");
        Integer port = Integer.parseInt(properties.getProperty("server.port"));
        connection = new SocketOpenOfficeConnection(host, port);
        documentConverter = new StreamOpenOfficeDocumentConverter(connection);
        try
        {
            connection.connect();
            if (isConnection())
            {
                return true;
            }
            return false;
        }
        catch (Exception e)
        {
            System.out.println("重连失败");
            e.printStackTrace();
            return false;
        }
    }

    /**
     * 得到当前properties
     *
     * @return properties对象
     */
    public static Properties getProperties()
    {
        return properties;
    }

    /**
     * 判断服务是否正常连接
     *
     * @return 成功连接为true
     */
    public static boolean isConnection()
    {
        if (connection == null || !connection.isConnected())
        {
            return false;
        }
        return true;
    }

    /**
     * 判断服务是否没有连接
     *
     * @return 成功连接为false
     */
    public static boolean isNotConnection()
    {
        return !isConnection();
    }


    /**
     * 文档转换
     *
     * @param sourcePath 源路径
     * @param targetPath 目标路径
     */
    public static void converter(String sourcePath, String targetPath)
    {
        documentConverter.convert(new File(sourcePath), new File(targetPath));
    }

    /**
     * 文档转换
     *
     * @param sourcePath 源路径
     * @param targetPath 目标路径
     */
    public static void converter(File sourcePath, File targetPath)
    {
        documentConverter.convert(sourcePath, targetPath);
    }

    /**
     * 文档转换
     *
     * @param inputStream      输入流
     * @param outputStream     输出流
     * @param sourceFileSuffix 源文件后缀
     * @param targetFileSuffix 目标文件后缀
     */
    public static void converter(InputStream inputStream, OutputStream outputStream,
                                 String sourceFileSuffix, String targetFileSuffix)
    {
        DocumentFormat sourceDocumentFormat = new DefaultDocumentFormatRegistry()
                .getFormatByFileExtension(sourceFileSuffix);
        DocumentFormat targetDocumentFormat = new DefaultDocumentFormatRegistry()
                .getFormatByFileExtension(targetFileSuffix);
        documentConverter.convert(inputStream, sourceDocumentFormat, outputStream, targetDocumentFormat);
    }
}
```





### 第四步：编写openoffice.properties

```properties
# openoffice服务所在地址
server.host=127.0.0.1
# openoffice服务端口号
server.port=8100
# 服务连接检查周期，单位是毫秒，600000为10分钟。如果服务连接挂了，将会尝试重新连接
serviceCheckCycle=600000
```



配置文件位置：

![image-20231125001548442](img/openoffice学习笔记/image-20231125001548442.png)





### 第五步：编写测试类

```java
package mao;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Project name(项目名称)：openoffice-word-to-pdf2
 * Package(包名): mao
 * Class(类名): Test
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/11/25
 * Time(创建时间)： 0:05
 * Version(版本): 1.0
 * Description(描述)： 无
 */

public class Test
{
    public static void main(String[] args) throws FileNotFoundException
    {
        DocumentConverterUtil.converter("./test.docx", "./test.pdf");
        DocumentConverterUtil.converter("./test.docx", "./test2.pdf");
        DocumentConverterUtil.converter("./test.docx", "./test3.pdf");
        DocumentConverterUtil.converter(new FileInputStream("./test.docx"),
                new FileOutputStream("./test4.pdf"),"docx","pdf");
        System.out.println("转换完成");
    }
}
```





### 第六步：运行查看结果

```sh
十一月 25, 2023 12:17:56 上午 com.artofsolving.jodconverter.openoffice.connection.AbstractOpenOfficeConnection connect
信息: connected
转换完成
```



源文档内容：

![image-20231125002042440](img/openoffice学习笔记/image-20231125002042440.png)





转换后：

![image-20231125002058289](img/openoffice学习笔记/image-20231125002058289.png)



结果正确，性能在1秒内









