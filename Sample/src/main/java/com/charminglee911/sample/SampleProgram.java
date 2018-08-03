package com.charminglee911.sample;

import com.charminglee911.excelutils.ExcelUtile;

import java.io.FileInputStream;
import java.util.HashMap;
import java.util.List;

/**
 * Created by CharmingLee on 2017/4/3.
 */
public class SampleProgram {

    public static void main(String[] args) throws Exception {
        String path = SampleProgram.class.getClassLoader().getResource("").getPath();
        String filePath1 = path + "person.xlsx";
        FileInputStream fileIS1 = new FileInputStream(filePath1);
        List<Person> people1 = ExcelUtile.xlsxToObj(fileIS1, Person.class);
        for (int i = 0; i < people1.size(); i++) {
            System.out.println(people1.get(i));
        }

        System.out.println("=====================");

        HashMap<String, String> map = new HashMap<String, String>();
        map.put("姓名","name");
        map.put("年龄","age");
        String filePath2 = path + "人物.xlsx";
        FileInputStream fileIS2 = new FileInputStream(filePath2);
        List<Person> people2 = ExcelUtile.xlsxToObj(fileIS2, Person.class, map);
        for (int i = 0; i < people2.size(); i++) {
            System.out.println(people2.get(i));
        }

        System.out.println("=====================");

        FileInputStream fileIS3 = new FileInputStream(filePath2);
        List<PersonAnnotation> people3 = ExcelUtile.xlsxToObj(fileIS3, PersonAnnotation.class, map);
        for (int i = 0; i < people3.size(); i++) {
            System.out.println(people3.get(i));
        }
    }

}
