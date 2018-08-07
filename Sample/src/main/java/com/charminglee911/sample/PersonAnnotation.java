package com.charminglee911.sample;

import com.charminglee911.excelutils.ExcelField;
import com.charminglee911.excelutils.ExcelSheet;

/**
 * @Author CharmingLee
 * @Date 2018/8/1
 * @Description TODO
 **/
@ExcelSheet(name = "person")
public class PersonAnnotation {
    @ExcelField(name = "姓名")
    private String name;
    @ExcelField(name = "年龄")
    private String age;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    @Override
    public String toString() {
        return "PersonAnnotation{" +
                "name='" + name + '\'' +
                ", age='" + age + '\'' +
                '}';
    }
}
