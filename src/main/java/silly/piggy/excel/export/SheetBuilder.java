package silly.piggy.excel.export;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import silly.piggy.excel.export.common.CommonBuilder;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Objects;
import java.util.function.Function;

public class SheetBuilder {

    private String sheetName;
    private ExcelBuilder excelBuilder;
    private Collection<?> sheetData;
    private List<CommonBuilder> commonBuilders = new ArrayList<>();

    protected SheetBuilder(String sheetName, ExcelBuilder excelBuilder) {
        this.sheetName = sheetName;
        this.excelBuilder = excelBuilder;
    }

    public SheetBuilder setSheetData(Collection<?> sheetData) {
        this.sheetData = sheetData;
        return this;
    }

    public <T> CommonBuilder commonColumn(char column, Function<T, Object> func) {
        CommonBuilder commonBuilder = new CommonBuilder(this, column, func);
        commonBuilder.setData(sheetData);
        this.commonBuilders.add(commonBuilder);
        return commonBuilder;
    }

    public ExcelBuilder next() {
        return excelBuilder;
    }

    protected void build(HSSFWorkbook workbook) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        HSSFSheet excelSheet = workbook.getSheet(this.sheetName);
        if (Objects.isNull(excelSheet)) {
            excelSheet = workbook.createSheet(this.sheetName);
        }

        for (CommonBuilder commonBuilder : commonBuilders) {
            Class clazz = commonBuilder.getClass();
            Method method = clazz.getDeclaredMethod("build", HSSFSheet.class);
            method.setAccessible(true);
            method.invoke(commonBuilder, excelSheet);
        }
    }

}
