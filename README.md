# ExcelToObject
数据从Excel表格转成对象bean

假前开会确认了新需求，我分分到的需求中有一个读取Excel表格的数据，找了github也看到有好用的工具类，毕竟Excel不像json这么热。假期时间宽裕，正好又好久不写博客了，干脆就写一个从Excel中的数据反向生成实体对象的小工具，思想继承自fastjson这类json转换工具，下面看Excel表

第一种是中文的属性名

![1936229-caa594111d4c0b76.png](http://upload-images.jianshu.io/upload_images/1936229-1a5500843531fcfc.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

第二种是英文网的属性名

![1936229-9e132f2be5426589.png](http://upload-images.jianshu.io/upload_images/1936229-3d237d2024653471.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

假设对象为Person

![1936229-b30b12cf94bdd8b0.png](http://upload-images.jianshu.io/upload_images/1936229-78790fab5df657ee.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

分别对2中情况的转换，操作很简单，思想继承自fastjson

![1936229-69262f702e529880.png](http://upload-images.jianshu.io/upload_images/1936229-55ec0a24cd71f9c5.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

2018.8.5日更新：
增加注解操作，具体的例子可以从我的github工程里看，截图操作如下：
![809056E9-2BEB-450F-B3D9-FE5B66D208F7.png](https://upload-images.jianshu.io/upload_images/1936229-1a7e34e68f43075a.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

后续更新导出Excel功能！！！

maven地址点[这里](https://mvnrepository.com/artifact/com.charminglee911/ExcelUtils)

maven构建：
```
<dependency>
    <groupId>com.charminglee911</groupId>
    <artifactId>ExcelUtils</artifactId>
    <version>0.9.release</version>
</dependency>
```

简书地址：
http://www.jianshu.com/p/5696317fd4c7
