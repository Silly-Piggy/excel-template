package silly.piggy;

import silly.piggy.excel.export.ExcelBuilder;

import java.lang.reflect.InvocationTargetException;
import java.util.Arrays;
import java.util.List;

public class AppTest {

    static List<Student> students_en = Arrays.asList(new Student("Lily", 12), new Student("Lucy", 23));
    static List<Student> students_cn = Arrays.asList(new Student("李雷", 12), new Student("韩梅梅", 23));


    public static void main(String[] args) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        test2();
    }

    static void test1() throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        new ExcelBuilder()
                .sheet("sheet_en").setSheetData(students_en)
                .commonColumn('A', Student::getName).next()
                .commonColumn('B', Student::getAge).next()
                .next()
                .saveToDisk("/var/root/Desktop/excel.xls");
    }

    static void test2() throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        new ExcelBuilder()
                .sheet("sheet_cn")
                .commonColumn('A', Student::getName).fromRow(1).setData(students_cn).next()
                .commonColumn('B', Student::getName).setData(students_en).next()
                .next()
                .saveToDisk("/var/root/Desktop/excel.xls");
    }

}
