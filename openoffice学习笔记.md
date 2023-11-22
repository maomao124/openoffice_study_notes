[toc]

---













# 概述













# openoffice

## 概述









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





![image-20231122172733550](img/openoffice学习笔记/image-20231122172733550.png)



![image-20231122172745153](img/openoffice学习笔记/image-20231122172745153.png)



![image-20231122172803011](img/openoffice学习笔记/image-20231122172803011.png)



![image-20231122172851825](img/openoffice学习笔记/image-20231122172851825.png)









### Linux安装











### Docker安装

















# SpringBoot整合openoffice

## 概述

因业务需要，需要使用SpringBoot将docx文件转换成PDF文件，再将PDF文件响应给用户



## 使用

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





























# java调用openoffice

## 概述







