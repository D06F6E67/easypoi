package cn.afterturn.easypoi.excel.export.template;

import cn.afterturn.easypoi.excel.entity.TemplateSumEntity;
import cn.afterturn.easypoi.util.PoiCellUtil;
import lombok.Getter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static cn.afterturn.easypoi.util.PoiElUtil.*;

/**
 * 针对模板合并同类项做统一处理
 * 1.处理模板之前统计需要确定合并同类项的关联关系
 * 2.遍历时获取关联的key以及col
 * 3.数据处理完后，在sum之前合并同类项
 */
public class TemplateMergeColHandler {

    // 列的key和下标
    private Map<String, Integer> keyMap = new HashMap<>();
    // 同类项依赖下标关系
    private Map<Integer, int[]> mergeMap = new HashMap<>();
    // 同类项依赖key关系
    private Map<Integer, String[]> mergeColMap = new HashMap<>();
    // 表头的行数
    @Getter
    private int startRow = 0;

    public TemplateMergeColHandler(Sheet sheet) {
        getAllMergeColCell(sheet);
    }

    /**
     * 处理所有需要合并同类项的单元格
     */
    private void getAllMergeColCell(Sheet sheet) {
        Row row;
        int index = 0;
        while (index <= sheet.getLastRowNum()) {
            row = sheet.getRow(index++);
            if (row == null) {
                continue;
            }
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                if (row.getCell(i) != null && PoiCellUtil.getCellValue(row.getCell(i)).contains(MERGE_COL)) {
                    addMergeColCellToList(row.getCell(i));
                    if (startRow == 0) {
                        startRow = row.getRowNum();
                    }
                }
            }
        }
    }

    private void addMergeColCellToList(Cell cell) {
        String mergeColKye = getMergeColKye(cell);
        if (StringUtils.isEmpty(mergeColKye)) {
            mergeMap.put(cell.getColumnIndex(), null);
        } else {
            mergeColMap.put(cell.getColumnIndex(), mergeColKye.split(","));
        }
    }

    /**
     * 获取key merge_col:(key)
     *
     * @param cell
     *
     * @return
     */
    private String getMergeColKye(Cell cell) {
        String cellValue = cell.getStringCellValue();
        int startIndex = cellValue.indexOf(MERGE_COL);
        String mergeCol = cellValue.substring(startIndex, cellValue.indexOf(RIGHT_BRACKET, startIndex) + 1);
        delMergeCol(cell, mergeCol);

        return mergeCol.substring(MERGE_COL.length() + 1, mergeCol.length() - 1);
    }

    /**
     * 删除单元格内mergeCol字符
     *
     * @param cell     单元格
     * @param mergeCol 字符
     */
    private void delMergeCol(Cell cell, String mergeCol) {
        cell.setCellValue(cell.getStringCellValue().replace(mergeCol, "").trim());
    }

    public void addKey(String key, Cell cell) {
        keyMap.put(key, cell.getColumnIndex());
    }

    public Map<Integer, int[]> getDataList() {

        if (keyMap.isEmpty()) {
            return mergeMap;
        }

        mergeColMap.forEach((key, value) -> {
            int[] rely = new int[value.length];
            for (int i = 0; i < value.length; i++) {
                rely[i] = keyMap.get(value[i]);
            }
            mergeMap.put(key, rely);
        });
        return mergeMap;
    }
}