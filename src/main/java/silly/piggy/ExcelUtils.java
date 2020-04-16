package silly.piggy;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ExcelUtils {

    private ExcelUtils(){}

    public static void addValidationToSheet(Workbook workbook, Sheet targetSheet, Map<String, List<String>> options, char keyColumn, char valueColumn, int fromRow, int endRow) {
        String hiddenSheetName = "sheet" + workbook.getNumberOfSheets();
        Sheet hiddenSheet = workbook.createSheet(hiddenSheetName);
        List<String> firstLevelItems = new ArrayList<>();

        int rowIndex = 0;
        for (Map.Entry<String, List<String>> entry : options.entrySet()) {
            String parent = formatNameName(entry.getKey());
            firstLevelItems.add(parent);
            List<String> children = entry.getValue();

            if (CollectionUtils.isEmpty(children)) {
                continue;
            }

            int columnIndex = 0;
            Row row = hiddenSheet.createRow(rowIndex++);
            Cell cell = null;

            for (String child : children) {
                cell = row.createCell(columnIndex++);
                cell.setCellValue(child);
            }

            char lastChildrenColumn = (char) ((int) 'A' + children.size() - 1);
            createName(workbook, parent, String.format("%s!$A$%s:$%s$%s", hiddenSheetName, rowIndex, lastChildrenColumn, rowIndex));

            DVConstraint constraint = DVConstraint.createFormulaListConstraint("INDIRECT($" + keyColumn + "1)");
            CellRangeAddressList regions = new CellRangeAddressList(fromRow, endRow, valueColumn - 'A', valueColumn - 'A');
            targetSheet.addValidationData(new HSSFDataValidation(regions, constraint));
        }

        addValidationToSheet(workbook, targetSheet, firstLevelItems.toArray(), keyColumn, fromRow, endRow);

    }

    /**
     * 给sheet页，添加下拉列表
     *
     * @param workbook    excel文件，用于添加Name
     * @param targetSheet 级联列表所在sheet页
     * @param options     级联数据 ['百度','阿里巴巴']
     * @param column      下拉列表所在列 从'A'开始
     * @param fromRow     级联限制开始行
     * @param endRow      级联限制结束行
     */
    public static void addValidationToSheet(Workbook workbook, Sheet targetSheet, Object[] options, char column, int fromRow, int endRow) {
        String hiddenSheetName = "sheet" + workbook.getNumberOfSheets();
        Sheet optionsSheet = workbook.createSheet(hiddenSheetName);
        String nameName = column + "_parent" + System.currentTimeMillis();

        int rowIndex = 0;
        for (Object option : options) {
            Row row = optionsSheet.createRow(rowIndex++);
            Cell cell = row.createCell(0);
            cell.setCellValue(option.toString());
        }

        createName(workbook, nameName, hiddenSheetName + "!$A$1:$A$" + options.length);

        DVConstraint constraint = DVConstraint.createFormulaListConstraint(nameName);
        CellRangeAddressList regions = new CellRangeAddressList(fromRow, endRow, (int) column - 'A', (int) column - 'A');
        targetSheet.addValidationData(new HSSFDataValidation(regions, constraint));
    }

    /**
     * 根据用户在keyColumn选择的key, 自动填充value到valueColumn
     *
     * @param workbook    excel
     * @param targetSheet sheet页
     * @param keyValues   匹配关系 {'百度','www.baidu.com'},{'淘宝','www.taobao.com'}
     * @param keyColumn   要匹配的列（例如 网站中文名称）
     * @param valueColumn 匹配到的内容列（例如 网址）
     * @param fromRow     下拉限制开始行
     * @param endRow      下拉限制结束行
     */
    public static void addAutoMatchValidationToSheet(Workbook workbook, Sheet targetSheet, Map<String, String> keyValues, char keyColumn, char valueColumn, int fromRow, int endRow) {
        String hiddenSheetName = "sheet" + workbook.getNumberOfSheets();
        Sheet hiddenSheet = workbook.createSheet(hiddenSheetName);

        // init the search region(A and B columns in hiddenSheet)
        int rowIndex = 0;
        for (Map.Entry<String, String> kv : keyValues.entrySet()) {
            Row totalSheetRow = hiddenSheet.createRow(rowIndex++);

            Cell cell = totalSheetRow.createCell(0);
            cell.setCellValue(kv.getKey());

            cell = totalSheetRow.createCell(1);
            cell.setCellValue(kv.getValue());
        }

        for (int i = fromRow; i <= endRow; i++) {
            Row totalSheetRow = targetSheet.getRow(i);
            if (totalSheetRow == null) {
                totalSheetRow = targetSheet.createRow(i);
            }

            Cell cell = totalSheetRow.getCell((int) valueColumn - 'A');
            if (cell == null) {
                cell = totalSheetRow.createCell((int) valueColumn - 'A');
            }

            String keyCell = String.valueOf(keyColumn) + (i + 1);
            String formula = String.format("IF(ISNA(VLOOKUP(%s,%s!A:B,2,0)),\"\",VLOOKUP(%s,%s!A:B,2,0))", keyCell, hiddenSheetName, keyCell, hiddenSheetName);

            cell.setCellFormula(formula);
        }

        // init the keyColumn as comboList
        addValidationToSheet(workbook, targetSheet, keyValues.keySet().toArray(), keyColumn, fromRow, endRow);
    }

    private static Name createName(Workbook workbook, String nameName, String formula) {
        Name name = workbook.createName();
        name.setNameName(nameName);
        name.setRefersToFormula(formula);
        return name;
    }

    /**
     * 不可数字开头
     *
     * @param name
     * @return
     */
    public static String formatNameName(String name) {
        name = name.replaceAll(" ", "").replaceAll("-", "_").replaceAll(":", ".").replaceAll("\\(|\\)", "").replaceAll("/", "").replaceAll("（", "").replaceAll("）", "").replaceAll("	", "").replaceAll("：", "");
        name = "_" + name;

        return name;
    }

    public static CellStyle createCellStyle(Workbook workbook, String fontName, int fontSize, boolean bold) {
        CellStyle headStyle = workbook.createCellStyle();

        Font font = workbook.createFont();
        font.setFontName(fontName);
        font.setBold(bold);
        font.setFontHeightInPoints((short) fontSize);
        headStyle.setFont(font);

        return headStyle;
    }

    public static CellStyle createHeaderStyle(Workbook workbook) {
        return createCellStyle(workbook, "黑体", 14, true);
    }

    public static CellStyle createContentStyle(Workbook workbook) {
        return createCellStyle(workbook, "黑体", 10, false);
    }

    public static Sheet createSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.createSheet(sheetName);
        sheet.setDefaultColumnWidth(13);
        return sheet;
    }

    public static void createHeaderCells(Sheet sheet, Integer rowIndex, CellStyle style, String[] headers) {
        Row totalSheetRow = sheet.createRow(rowIndex);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = totalSheetRow.createCell(i);
            cell.setCellStyle(style);
            cell.setCellValue(headers[i]);
        }
    }

    public static String valueOf(Object obj) {
        if (obj == null) {
            return "";
        }
        return obj.toString();
    }
}
