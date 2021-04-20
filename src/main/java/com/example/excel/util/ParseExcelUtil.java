package com.example.excel.util;
import com.example.excel.pojo.TableHeadEnum;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.InputStream;
import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * @author liujunhui
 * @date 2020/9/7 20:02
 */
public class ParseExcelUtil {

    /**
     * 解析成绩Excel文件数据
     * @param file
     * @return
     */
    public static List<Map<String, Object>> getExcelData(InputStream file, Boolean transcolumnsQualifier) {
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(file);
        } catch (Exception e) {
            System.out.println("创建文档失败，IO异常信息如下");
            return null;
        }
        
        // 获取sheet的数量
        int sheetCount = workbook.getNumberOfSheets();
        if (sheetCount > 1) {
            System.out.println("excel含有多个sheet,仅解析第一个的内容");
        }
        List<Map<String, Object>> result = parseExcelSheet(workbook.getSheetAt(0), false);
        if (!CollectionUtils.isEmpty(result)) {
            // map中的所有键值对的值均为空，不处理该行数据
            return result.stream().filter(map -> !isEmptyMap(map)).collect(Collectors.toList());
        } else {
            return null;
        }
    }

    /**
     * 解析学生成绩excel文件数据
     * @param sheet                 处理的表单对象
     * @param transColumnsQualifier map中key值是否转换为hbase表中字段名
     * @return
     */
    public static List<Map<String, Object>> parseExcelSheet(Sheet sheet, boolean transColumnsQualifier) {

        List<Map<String, Object>> outPutData = new ArrayList<>();

        // 最大行数
        int rowNum = sheet.getLastRowNum();

        // 1.excel所有数据
        Map<Integer, List<Object>> rowMap = new LinkedHashMap<>();
        for (int i = 0; i <= rowNum; i++) {
            Row rowTemp = sheet.getRow(i);
            if (isBlackRow(rowTemp)) {
                continue;
            }
            List<Object> rowCellValues = new ArrayList<>();
            for (int j = 0; j < rowTemp.getLastCellNum(); j++) {
                if (isMergedRegion(sheet, i, j)) {
                    rowCellValues.add(getMergedRegionValue(sheet, i, j));
                } else {
                    rowCellValues.add(getCellFormatValue(rowTemp.getCell(j)).toString());
                }
            }
            rowMap.put(i, rowCellValues);
        }

        // 2.表头行数据
        Map<Integer, List<Object>> headRowMap = rowMap.entrySet().stream()
                .filter(entry -> entry.getValue().contains(TableHeadEnum.SCORE_HEAD_4.getHeadName()))
                .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue));
        // fixme 没有找到“姓名”行，取第一行作为表头
        if (headRowMap.size() == 0) {
            headRowMap = rowMap.entrySet().stream().filter(entry -> entry.getKey() == 0)
                    .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue));
        }
        for (Map.Entry<Integer, List<Object>> entry : headRowMap.entrySet()) {
            headRowMap.put(entry.getKey(), deleteRowBlankTail(entry.getValue()));
        }

        // 表头开始行、结束行的行数，表头列数
        int headRowStartIndex = Collections.min(headRowMap.keySet());
        int headRowEndIndex = Collections.max(headRowMap.keySet());
        int headRowCellNum = headRowMap.get(headRowStartIndex).size();

        // 表头内容
        List<String> headNameList = new ArrayList<>();
        for (int j = 0; j < headRowCellNum; j++) {
            Set<String> tempHead = new LinkedHashSet<>();
            for (int i = headRowStartIndex; i <= headRowEndIndex; i++) {
                try {
                    tempHead.add(headRowMap.get(i).get(j).toString());
                } catch (IndexOutOfBoundsException e) {
                    tempHead.add("");
                }
            }
            String headName = StringUtils.deleteWhitespace(tempHead.stream().collect(Collectors.joining()));
            if (transColumnsQualifier) {
                headNameList.add(TableHeadEnum.getByHeadName(headName).getColumnQualifier());
            } else {
                headNameList.add(j + "-" + headName);
            }
        }

        // 3.表内容行数据
        Map<Integer, List<Object>> contentRowMap = rowMap.entrySet().stream().filter(map -> map.getKey() > headRowEndIndex)
                .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue));

        for (Map.Entry<Integer, List<Object>> entry : contentRowMap.entrySet()) {
            Map<String, Object> dataMap = new LinkedHashMap<>();

            for (int i = 0; i < headNameList.size(); i++) {
                String headName = headNameList.get(i).substring(headNameList.get(i).indexOf("-") + 1);
                if (TableHeadEnum.HEAD_UNKNOWN.getHeadName().equals(headName)) {
                    continue;
                }
                try {
                    dataMap.put(headName, entry.getValue().get(i));
                } catch (IndexOutOfBoundsException e) {
                    dataMap.put(headName, "");
                }
            }
            outPutData.add(dataMap);
        }
        return outPutData;
    }


    /**
     * 判断单元格是否为合并单元格
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    private static boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }


    /**
     * 获取合并单元格的值
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int laseRow = range.getLastRow();
            if (row >= firstRow && row <= laseRow
                    && column >= firstColumn && column <= lastColumn) {
                Cell firstCell = sheet.getRow(firstRow).getCell(firstColumn);
                return String.valueOf(firstCell);
            }
        }
        return null;
    }


    /**
     * 获取单元格的值
     * @param cell
     * @return
     */
    public static Object getCellFormatValue(Cell cell){
        Object cellValue = null;
        if (cell != null){
            // 判断cell类型
            switch (cell.getCellType()) {
                case NUMERIC: {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    if (Pattern.compile(".*E[0-9]+.*").matcher(cellValue.toString()).matches()) {
                        cellValue = cellValue.toString().substring(0, cellValue.toString().indexOf("E"))
                                .replace(".", "");
                    }
                    break;
                }
                case FORMULA: {
                    cellValue = cell.getCellFormula();
                    break;
                }
                case STRING: {
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                case BOOLEAN: {
                    cellValue = cell.getBooleanCellValue();
                    break;
                }
                default:
                    cellValue = "";
            }
        } else {
            cellValue = "";
        }
        if (cellValue.toString().endsWith(".0")) {
            cellValue = cellValue.toString().substring(0, cellValue.toString().length() - 2);
        }
        return cellValue;
    }


    /**
     * 判断是否是空行
     * @param row
     * @return
     */
    public static boolean isBlackRow(Row row) {
        if (row == null || row.getLastCellNum() <= 0) {
            return true;
        }
        int blackCellNum = 0;
        for (int i = 0;i < row.getLastCellNum(); i++) {
            blackCellNum = StringUtils.isBlank(getCellFormatValue(row.getCell(i)).toString()) ? ++blackCellNum : blackCellNum;
        }
        return blackCellNum == row.getLastCellNum() ? true : false;
    }


    /**
     * 删除表头行后面为空值的尾部
     * @param headNameList
     * @return
     */
    private static List<Object> deleteRowBlankTail(List<Object> headNameList) {
        List<Object> temp = headNameList;
        for (int i = temp.size() - 1; i >= 0; i--) {
            String headName = headNameList.get(i).toString();
            if (StringUtils.isNotBlank(headName)) {
                temp = temp.subList(0, i + 1);
                break;
            }
        }
        return temp;
    }


    /**
     * 获取Excel文件第一个表单的第一行数据
     * @param file
     * @return
     */
    public static String getExcelTitle(InputStream file) {
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(file);
        } catch (Exception e) {
            System.out.println("创建文档失败，IO异常信息如下");
            return null;
        }
        // 第一行数据
        Row row = workbook.getSheetAt(0).getRow(0);
        if (row == null) {
            System.out.println("表中数据为空，请检查！");
            return null;
        }
        StringBuilder stringBuilder = new StringBuilder();
        for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
            stringBuilder.append(row.getCell(i).toString());
        }

        return stringBuilder.toString();
    }
    

    /**
     * 判断Map键值对的key对应的值是否全为空
     * @param map
     * @return
     */
    private static boolean isEmptyMap(Map<String, Object> map) {
        int result = 0;
        for (Map.Entry<String, Object> entry : map.entrySet()) {
            if (entry.getValue() == null || StringUtils.isBlank(entry.getValue().toString())) {
                result = result + 1;
            }
        }
        return result == map.size();
    }
}
