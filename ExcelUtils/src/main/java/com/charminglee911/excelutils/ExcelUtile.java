package com.charminglee911.excelutils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * Author CharmingLee
 * Date 2017/4/1
 * Description 将Excel表中的内容生成对应的对象，第0行必须是对象的属性命，如果第0行和对象属性名不一致，可进行相关的关系映射
 *
 */
public class ExcelUtile {
    /**
     * Excel表中属性集合
     */
    private static List<String> fieldList = new ArrayList<String>();
    /**
     * Excel表中属性和对象的关系映射
     */
    private static Map<String, String> fieldMapped = new  HashMap<String, String>();

    private static final String LANG_STRING = "java.lang.String";
    private static final String LANG_INTEGER = "java.lang.Integer";
    private static final String LANG_DOUBLE = "java.lang.Double";
    private static final String LANG_SHORT = "java.lang.Short";
    private static final String LANG_LONG = "java.lang.Long";
    private static final String LANG_FLOAT = "java.lang.Float";
    private static final String LANG_BOOLEAN = "java.lang.Boolean";


    /**
     * 2007以下版本生成的xls格式的excle
     * @param xlsxIS    文件流
     * @param classe    要解析成的对象
     * @param <T>       泛型
     * @return
     * @throws Exception
     */
    public static<T> List<T> xlsToObj(InputStream xlsxIS, Class<T> classe) throws Exception{
        return xlsToObj(xlsxIS, classe, null);
    }

    /**
     * 2007以上版本生成的xls格式的excle
     * @param xlsxIS    文件流
     * @param classe    要解析成的对象
     * @param <T>       泛型
     * @return
     * @throws Exception
     */
    public static<T> List<T> xlsxToObj(InputStream xlsxIS, Class<T> classe) throws Exception{
        return xlsxToObj(xlsxIS, classe, null);
    }

    /**
     * 2007以下版本生成的xls格式的excle
     * @param xlsxIS    文件流
     * @param classe    要解析成的对象
     * @param mapped    对象属性和excle表中的字段的映射关系
     *                  key为Excel表中的字段，value为classe中对应的属性名称
     * @param <T>       泛型
     * @return
     * @throws Exception
     */
    public static<T> List<T> xlsToObj(InputStream xlsxIS, Class<T> classe, Map<String, String> mapped) throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook(xlsxIS);
        List<T> list = excelToObj(workbook, classe, mapped);

        return list;
    }

    /**
     * 2007以上版本生成的xlsx格式的excle
     * @param xlsxIS    文件流
     * @param classe    要解析成的对象
     * @param mapped    对象属性和excle表中的字段的映射关系
     *                  key为Excel表中的字段，value为classe中对应的属性名称
     * @param <T>       泛型
     * @return
     * @throws Exception
     */
    public static<T> List<T> xlsxToObj(InputStream xlsxIS, Class<T> classe, Map<String, String> mapped) throws Exception {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(xlsxIS);
        return excelToObj(xssfWorkbook, classe, mapped);
    }

    /**
     * 解析excle中的字段成对象
     * @param workbook  excle对象
     * @param classe    要解析成的对象
     * @param mapped    对象属性和excle表中的字段的映射关系
     *                  key为Excel表中的字段，value为classe中对应的属性名称
     * @param <T>       泛型
     * @return
     * @throws Exception
     */
    private static<T> List<T> excelToObj(Workbook workbook, Class<T> classe, Map<String, String> mapped) throws Exception {
        //创建对象集合
        List<T> list = new ArrayList<T>();

        //循环所有表格生成对象
        for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
            Sheet sheet = workbook.getSheetAt(numSheet);
            if (sheet == null)
                continue;

            //生成Excel表中的属性字段和对象属性的映射关系
            createFieldMapped(sheet, mapped, classe);

            //生成对象，并读取Excel表中的字段给对象设置相应属性，并添加到list中
            createObjs(list, sheet, classe);
        }

        fieldList = new ArrayList<String>();
        fieldMapped = new  HashMap<String, String>();

        return list;
    }

    /**
     * 生成Excel表中的属性字段和对象属性的映射关系
     * @param list
     * @param sheet
     * @param classe
     * @param <T>
     * @throws Exception
     */
    private static<T> void createObjs(List<T> list, Sheet sheet, Class<T> classe) throws Exception{

        //第0行默认为对象属性，从第1行读取字段作为对象的属性
        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null)
                continue;

            //创建一个对象
            T obj = classe.newInstance();
            list.add(obj);
            for (int i = 0 ; i < fieldList.size(); i++){
                //获取该列属性在对象中对应的属性
                String key = fieldList.get(i);
                key = fieldMapped.get(key);

                //excle表中的行
                Cell cell = row.getCell(i);

                //设置对象属性值
                setObjField(obj, classe, key, cell);
            }

        }

    }

    /**
     * 生成对象，并读取Excel表中的字段给对象设置相应属性，并添加到list中
     * @param sheet
     * @param mapped
     * @param classe
     */
    private static<T> void createFieldMapped(Sheet sheet, Map<String, String> mapped, Class<T> classe){
        //拿到第0行，每列默认为对象属性名
        Row fieldsRow = sheet.getRow(sheet.getFirstRowNum());
        if (fieldsRow == null){
            return;
        }

        //判断是否存在映射关系
        boolean isMapping = (mapped != null && !mapped.isEmpty());
        //判断是否存在注解映射
        boolean isAnnotation = isAnnotation(classe);

        for (short fieldIndex = fieldsRow.getFirstCellNum();
             fieldIndex < fieldsRow.getLastCellNum();
             fieldIndex++){

            Cell cell = fieldsRow.getCell(fieldIndex);
            String cellFiedl = getCellValue(cell);
            fieldList.add(cellFiedl);

            //处理对象属性和exle的映射
            if (isMapping){
                String value = mapped.get(cellFiedl);
                if (value != null && !"".equals(value)){
                    fieldMapped.put(cellFiedl, value);
                }else {
                    fieldMapped.put(cellFiedl, cellFiedl);
                }

            } else if (isAnnotation) {
                Field[] declaredFields = classe.getClass().getDeclaredFields();
                for (Field f : declaredFields) {
                    ExcelField annotation = f.getAnnotation(ExcelField.class);
                    if (annotation != null){
                        fieldMapped.put(cellFiedl, annotation.name());
                    }
                }
            } else { //没有映射关系，则默认使用表格中第0行作为对象的属性名
                fieldMapped.put(cellFiedl, cellFiedl);
            }

        }

    }

    /**
     * 判读是否注解映射
     * @param classe
     * @param <T>
     * @return
     */
    private static<T> boolean isAnnotation(Class<T> classe){
        boolean isTypeAnnotation = classe.getClass().isAnnotationPresent(ExcelSheet.class);
        if (isTypeAnnotation){
            return true;
        }

        Field[] declaredFields = classe.getClass().getDeclaredFields();
        for (Field f: declaredFields) {
            if (f.isAnnotationPresent(ExcelField.class)){
                return true;
            }
        }

        return false;
    }

    /**
     * 根据映射关系，给属性设置值
     * @param obj
     * @param classe
     * @param key
     * @param cell
     * @throws IllegalAccessException
     */
    private static void setObjField(Object obj, Class classe, String key, Cell cell) throws IllegalAccessException {
        //获取要设置的属性
        Field field = null;
        Field[] fields = classe.getDeclaredFields();
        for (Field f: fields) {
            if (f.getName().equals(key)){
                field = f;
                break;
            }
        }

        if (field == null)
            return;

        //判断field类型
        Object value = convertValue(field, cell);

        //设置属性
        field.setAccessible(true);
        field.set(obj, value);
    }

    /**
     * 把cell的value转换成和对象一样的属性类型
     * @param field
     * @param cell
     * @return
     */
    private static Object convertValue(Field field, Cell cell){
        String type = field.getType().getName();
//        String cellValue = getCellValue(cell);

        if (LANG_STRING.equals(type)){
            return getCellValue(cell);
        }

        if ("int".equals(type) || LANG_INTEGER.equals(type)){
            Integer integer = 0;
            try {
                String cellValue = getCellValue(cell);
                double dValue = Double.valueOf(cellValue);
                if (dValue % 1 == 0)
                    integer = Integer.valueOf((int) dValue);
            } catch (NumberFormatException e) {
                e.printStackTrace();
            }

            return integer;
        }

        if ("double".equals(type) || LANG_DOUBLE.equals(type)){
            Double aDouble = 0.0;
            try {
                aDouble = Double.valueOf(getCellValue(cell));
            } catch (NumberFormatException e) {
                e.printStackTrace();
            }

            return aDouble;
        }

        if ("short".equals(type) || LANG_SHORT.equals(type)){
            Short value = 0;
            try {
                String cellValue = getCellValue(cell);
                double dValue = Double.valueOf(cellValue);
                if (dValue % 1 == 0)
                    value = Short.valueOf((short) dValue);
            } catch (NumberFormatException e) {
                e.printStackTrace();
            }

            return value;
        }

        if ("long".equals(type) || LANG_LONG.equals(type)){
            Long value = 0L;
            try {
                String cellValue = getCellValue(cell);
                double dValue = Double.valueOf(cellValue);
                if (dValue % 1 == 0)
                    value = Long.valueOf((long) dValue);
            } catch (NumberFormatException e) {
                e.printStackTrace();
            }

            return value;
        }

        if ("float".equals(type) || LANG_FLOAT.equals(type)){
            Float value = 0F;
            try {
                value = Float.valueOf(getCellValue(cell));
            } catch (NumberFormatException e) {
                e.printStackTrace();
            }

            return value;
        }

        if ("boolean".equals(type) || LANG_BOOLEAN.equals(type)){
            Boolean value = false;
            try {
                value = Boolean.valueOf(getCellValue(cell));
            } catch (NumberFormatException e) {
                e.printStackTrace();
            }

            return value;
        }

        if ("char".equals(type)){
            char value = 0;
            try {
                String sValue = String.valueOf(getCellValue(cell));
                if (sValue.length() > 0)
                    value = sValue.charAt(0);
            } catch (NumberFormatException e) {
                e.printStackTrace();
            }

            return value;
        }

        return null;
    }

    /**
     * 从cell中获取Str值
     * @param cell
     * @return
     */
    private static String getCellValue(Cell cell){
        if (cell == null)
            return "";

        if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN)
            return String.valueOf(cell.getBooleanCellValue());

        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
            return String.valueOf(cell.getNumericCellValue());

        if (cell.getCellType() == Cell.CELL_TYPE_STRING)
            return cell.getStringCellValue();

        return "";
    }

}

