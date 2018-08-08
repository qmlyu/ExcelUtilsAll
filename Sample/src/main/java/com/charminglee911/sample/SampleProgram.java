package com.charminglee911.sample;

import com.charminglee911.excelutils.ExcelUtile;

import java.io.File;
import java.util.HashMap;
import java.util.List;

/**
 * Created by CharmingLee on 2017/4/3.
 */
public class SampleProgram {

    public static void main(String[] args) {
        String path = SampleProgram.class.getClassLoader().getResource("").getPath();
        String filePath1 = path + "person.xlsx";
        File file1 = new File(filePath1);
        List<Person> people1 = ExcelUtile.excelFileToObjects(file1, Person.class);
        for (int i = 0; i < people1.size(); i++) {
            System.out.println(people1.get(i));
        }
        System.out.println("=====================");

        HashMap<String, String> map = new HashMap<String, String>();
        map.put("姓名","name");
        map.put("年龄","age");
        String filePath2 = path + "人物.xlsx";
        File file2 = new File(filePath2);
        List<Person> people2 = ExcelUtile.excelFileToObjects(file2, Person.class, map);
        for (int i = 0; i < people2.size(); i++) {
            System.out.println(people2.get(i));
        }
        System.out.println("=====================");

        File file3 = new File(filePath2);
        List<PersonAnnotation> people3 = ExcelUtile.excelFileToObjects(file3, PersonAnnotation.class);
        for (int i = 0; i < people3.size(); i++) {
            System.out.println(people3.get(i));
        }

        //对象导出Excel文件
        File file = new File("/Users/charminglee/Desktop/test.xlsx");
        ExcelUtile.objectsToExcelFile(file, people3);
    }

}
