package org.tdubuis.filedata;

import lombok.Data;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
@Data
public class ExcelData {
    private String region;
    private HashMap<String, Data> dataMapMonth;
    private HashMap<String, Data> dataMapYTD;


    public ExcelData(String region) {
        this.region = region;
    }




    @lombok.Data
    public static class Data {
        private List<List<XSSFCell>> excelCells = new ArrayList<>();
        private List<CellRangeAddress> mergedRegion = new ArrayList<>();
        private HashMap<Integer,Float> columnWidth = new HashMap<>();

        public void shiftMergedRegionToOrigin() {
            XSSFCell firstCell = excelCells.get(0).get(0);
            int rowShift = firstCell.getRowIndex();
            int columnShift = firstCell.getColumnIndex();
            for (CellRangeAddress rangeAddress : mergedRegion) {
                rangeAddress.setFirstRow(rangeAddress.getFirstRow() - rowShift);
                rangeAddress.setLastRow(rangeAddress.getLastRow() - rowShift);
                rangeAddress.setFirstColumn(rangeAddress.getFirstColumn() - columnShift);
                rangeAddress.setLastColumn(rangeAddress.getLastColumn() - columnShift);
            }
        }
    }
}
