package cn.afterturn.easypoi.excel.export.template;

import static cn.afterturn.easypoi.util.PoiElUtil.*;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import cn.afterturn.easypoi.excel.entity.TemplateSumEntity;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import cn.afterturn.easypoi.util.PoiCellUtil;

/**
 * 针对模板统计问题做统一处理
 * 1.处理模板之前统计需要SUM的数据以及位置
 * 2.遍历时统计数据
 * 3.遍历后设置数据
 * @author JueYue
 * 2016年6月19日
 */
public class TemplateSumHandler {

    private Map<String, TemplateSumEntity> sumMap = new HashMap<String, TemplateSumEntity>();

    public TemplateSumHandler(Sheet sheet) {
        getAllSumCell(sheet);
    }

    /**
     * 统计计算所有的统计单元格
     */
    private void getAllSumCell(Sheet sheet) {
        Row row = null;
        int index = 0;
        while (index <= sheet.getLastRowNum()) {
            row = sheet.getRow(index++);
            if (row == null) {
                continue;
            }
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                if (row.getCell(i) != null && PoiCellUtil.getCellValue(row.getCell(i)).contains(SUM)) {
                    addSumCellToList(row.getCell(i));
                }
            }
        }
    }

    private void addSumCellToList(Cell cell) {
        String cellValue = cell.getStringCellValue();
        int index = 0;
        while ((index = indexOfIgnoreCase(cellValue, SUM, index)) != -1) {
            TemplateSumEntity entity = new TemplateSumEntity();
            entity.setCellValue(cellValue);
            entity.setSumKey(getSumKey(cellValue, index++));
            entity.setCol(cell.getColumnIndex());
            entity.setRow(cell.getRowIndex());
            sumMap.put(entity.getSumKey(), entity);
        }
    }

    public boolean isSumKey(String key) {
        return sumMap.containsKey(key);
    }

    /**
     * SUM:(key)
     * 
     * @param cellValue
     * @param index 
     * @return
     */
    private String getSumKey(String cellValue, int index) {
        return cellValue.substring(index + 5, cellValue.indexOf(")", index));
    }

    public void addValueOfKey(String key, String val) {
        if (StringUtils.isNoneEmpty(key)) {
            sumMap.get(key).setValue(sumMap.get(key).getValue().add(new BigDecimal(val)));
        }
    }

    public List<TemplateSumEntity> getDataList() {
        return new ArrayList<TemplateSumEntity>(sumMap.values());
    }

    public void addListSizeToSumEntity() {

    }

    /**
     * 
     * @param rowIndex
     * @param size 
     */
    public void shiftRows(int rowIndex, int size) {
        for (TemplateSumEntity entity : getDataList()) {
            if (entity.getRow() >= rowIndex) {
                entity.setRow(entity.getRow() + size);
            }
        }

    }

    private static int indexOfIgnoreCase(String str, String searchStr, int startPos) {
        if (str == null || searchStr == null) {
            return -1;
        }
        if (startPos < 0) {
            startPos = 0;
        }
        int endLimit = (str.length() - searchStr.length()) + 1;
        if (startPos > endLimit) {
            return -1;
        }
        if (searchStr.length() == 0) {
            return startPos;
        }
        for (int i = startPos; i < endLimit; i++) {
            if (str.regionMatches(true, i, searchStr, 0, searchStr.length())) {
                return i;
            }
        }
        return -1;
    }
}