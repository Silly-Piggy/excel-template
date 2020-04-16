package silly.piggy.excel.export.common;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import silly.piggy.excel.export.SheetBuilder;

import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Objects;
import java.util.function.Function;

public class CommonBuilder<T> {

    SheetBuilder sheetBuilder;

    private char columnIndex;
    private Integer rowIndex = 0;
    private Collection<T> data;
    private Function<T, Object> func;

    public CommonBuilder(SheetBuilder sheetBuilder, char columnIndex, Function<T, Object> func) {
        this.sheetBuilder = sheetBuilder;
        this.columnIndex = columnIndex;
        this.func = func;
    }

    public CommonBuilder fromRow(int rowIndex) {
        this.rowIndex = rowIndex;
        return this;
    }

    public SheetBuilder next() {
        return sheetBuilder;
    }

    public CommonBuilder setData(Collection<T> data) {
        if (Objects.isNull(data)) {
            return this;
        }
        this.data = new ArrayList<>(data.size());
        this.data.addAll(data);
        return this;
    }

    void build(HSSFSheet excelSheet) {
        for (T item : data) {
            HSSFRow row = excelSheet.getRow(rowIndex);
            if (Objects.isNull(row)) {
                row = excelSheet.createRow(rowIndex);
            }
            Cell cell = row.createCell(this.columnIndex - 'A');
            cell.setCellValue(func.apply(item).toString());

            rowIndex++;
        }
    }
}
