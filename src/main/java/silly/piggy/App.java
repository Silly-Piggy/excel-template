package silly.piggy;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class App {
    public static void main(String[] args) throws IOException {
        File file = new File("/var/root/Desktop/excel.xls");
        FileOutputStream outputStream = new FileOutputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet excelSheet = workbook.createSheet("excel");

        //demo 单独下拉列表
        ExcelUtils.addValidationToSheet(workbook, excelSheet, new String[]{"百度", "阿里巴巴"}, 'C', 1, 200);

        //demo 级联下拉列表
        Map<String, List<String>> data = new HashMap<>();
        data.put("百度系列", Arrays.asList("百度地图", "百度知道", "百度音乐"));
        data.put("阿里系列", Arrays.asList("淘宝", "支付宝", "钉钉"));
        ExcelUtils.addValidationToSheet(workbook, excelSheet, data, 'A', 'B', 1, 200);

        //demo 自动填充
        Map<String, String> kvs = new HashMap<>();
        kvs.put("百度", "www.baidu.com");
        kvs.put("阿里", "www.taobao.com");
        ExcelUtils.addAutoMatchValidationToSheet(workbook, excelSheet, kvs, 'D', 'E', 1, 200);

        // 隐藏存储下拉列表数据的sheet；可以注释掉该行以便查看、理解存储格式
        //hideTempDataSheet(workbook, 1);

        workbook.write(outputStream);
        outputStream.close();

    }
}
