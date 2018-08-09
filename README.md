# ExcelToObject
数据从Excel表格转成javaBean,从javaBean导出成Excel.

maven地址点[这里](https://mvnrepository.com/artifact/com.charminglee911/ExcelUtils)

maven构建：
```
<dependency>
    <groupId>com.charminglee911</groupId>
    <artifactId>ExcelUtils</artifactId>
    <version>1.1.release</version>
</dependency>
```

2018.8.8 更新:
1.修改Excel导入对象时使用File参数
2.导出支持xls和xlsx两种格式
3.识别File的文件扩展名，用以决定使用哪种兼容格式
4.升级poi到3.1.7
5.注释老接口过期
6.bugfix

2018.8.7 更新:
1.导出Excel功能！！！

2018.8.5 更新：
1.增加注解操作，具体的例子可以从我的github工程里看，截图操作如下：
![809056E9-2BEB-450F-B3D9-FE5B66D208F7.png](https://upload-images.jianshu.io/upload_images/1936229-1a7e34e68f43075a.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

假前开会确认了新需求，我分到的需求中有一个读取Excel表格的数据，找了github也看到有好用的工具类，毕竟Excel不像json这么热。假期时间宽裕，正好又好久不写博客了，干脆就写一个从Excel中的数据反向生成实体对象的小工具，思想继承自fastjson这类json转换工具，下面看Excel表

第一种是中文的属性名

![1936229-caa594111d4c0b76.png](http://upload-images.jianshu.io/upload_images/1936229-1a5500843531fcfc.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

第二种是英文网的属性名

![1936229-9e132f2be5426589.png](http://upload-images.jianshu.io/upload_images/1936229-3d237d2024653471.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

假设对象为Person

![1936229-b30b12cf94bdd8b0.png](http://upload-images.jianshu.io/upload_images/1936229-78790fab5df657ee.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

项目功能演示
![屏幕快照 2018-08-09 08.04.54.png](https://upload-images.jianshu.io/upload_images/1936229-e7301610d88672e5.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

项目功能演示(旧版本)
![屏幕快照 2018-08-07 23.25.35.png](https://upload-images.jianshu.io/upload_images/1936229-9c0bd43d73560524.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

简书地址：
http://www.jianshu.com/p/5696317fd4c7
