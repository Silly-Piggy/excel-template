package silly.piggy.excel.export;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

public class ExcelBuilder {

    private HSSFWorkbook workbook;
    private List<SheetBuilder> sheetBuilderList = new ArrayList<>();

    public SheetBuilder sheet(String sheetName) {
        SheetBuilder sheetBuilder = new SheetBuilder(sheetName, this);
        sheetBuilderList.add(sheetBuilder);
        return sheetBuilder;
    }

    private ExcelBuilder build() throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        this.workbook = new HSSFWorkbook();
        for (SheetBuilder sheetBuilder : sheetBuilderList) {
            sheetBuilder.build(workbook);
        }
        return this;
    }

    public void saveToDisk(String filePath) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        File file = new File(filePath);
        try (FileOutputStream outputStream = new FileOutputStream(file)) {
            build();
            this.workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
