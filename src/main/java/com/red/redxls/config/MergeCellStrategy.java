package com.red.redxls.config;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.merge.AbstractMergeStrategy;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.util.List;

/**
 * @author pjh
 * @created 2024/7/23
 */

public class MergeCellStrategy extends AbstractMergeStrategy {
    @Override
    protected void merge(Sheet sheet, Cell cell, Head head, Integer relativeRowIndex) {
        if (relativeRowIndex == null || relativeRowIndex == 0) {
            return;
        }
        int rowIndex = cell.getRowIndex();
        int colIndex = cell.getColumnIndex();
        sheet = cell.getSheet();

        Row thisRow = sheet.getRow(rowIndex);
        Row preRow = sheet.getRow(rowIndex - 1);

        short firstCellNum = thisRow.getFirstCellNum();
        short lastCellNum = thisRow.getLastCellNum();
        short preFirstCellNum = preRow.getFirstCellNum();
        short preLastCellNum = preRow.getLastCellNum();

        /*
        此方法为编辑单个有占位符的单元格时触发。
        为了让前后没有占位符的单元格也能正常合并以及继承前边一行样式，需要对此行的第一个与最后一个有占位符的单元格进行特殊处理

         */
        // 第一个有占位符的单元格
        if(colIndex==firstCellNum){
            if(preFirstCellNum < colIndex){// 上一行的首个单元格比此行首个单元格靠前
                for (int i = preFirstCellNum; i < firstCellNum; i++) {
                    // 复制上一行的样式
                    copyStyle(sheet,rowIndex -1 ,i,rowIndex,i);
                    // 合并单元格
                    if(getCellValue(thisRow.getCell(i)).equals(getCellValue(preRow.getCell(i)))){// 内容相同，合并
                        addMergeRange(sheet,rowIndex -1 ,rowIndex,i,i);
                    }
                }
            }

        }
        // 最后一个有占位符的单元格
        if(colIndex == lastCellNum){
            if(preLastCellNum > colIndex){// 上一行的最后一个单元格比此行最后一个单元格靠后
                for (int i = (lastCellNum+1); i <= preLastCellNum; i++) {
                    // 复制上一行的样式
                    copyStyle(sheet,rowIndex -1 ,i,rowIndex,i);
                    // 合并单元格
                    if(getCellValue(thisRow.getCell(i)).equals(getCellValue(preRow.getCell(i)))){// 内容相同，合并
                        addMergeRange(sheet,rowIndex -1,rowIndex ,i,i);
                    }
                }
            }
        }
        // 其他有占位符的单元格，只需处理自身即可
        if(colIndex >= firstCellNum && colIndex <= lastCellNum){
            // 复制上一行的样式
            copyStyle(sheet,rowIndex -1 ,colIndex,rowIndex,colIndex);
            if(getCellValue(thisRow.getCell(colIndex)).equals(getCellValue(preRow.getCell(colIndex)))){// 内容相同，合并
                addMergeRange(sheet,rowIndex -1 ,rowIndex,colIndex,colIndex);
            }
        }

    }

    private static void addMergeRange(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
    {
        CellRangeAddress cra = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        //在添加合并范围之前，检查已添加的合并范围，如果重叠，则删除原有合并范围，创建一个更大的。
        List<CellRangeAddress> list = sheet.getMergedRegions();
        for(CellRangeAddress cellRangeAddress : list){
            if(cellRangeAddress.intersects(cra)){

                int newFirstRow = Math.min(cellRangeAddress.getFirstRow(), firstRow);
                int newLastRow = Math.max(cellRangeAddress.getLastRow(), lastRow);
                int newFirstCol = Math.min(cellRangeAddress.getFirstColumn(), firstCol);
                int newLastCol = Math.max(cellRangeAddress.getLastColumn(), lastCol);

                cra = new CellRangeAddress(newFirstRow, newLastRow, newFirstCol,newLastCol);
                sheet.removeMergedRegion(list.indexOf(cellRangeAddress));
                break;
            }
        }
        sheet.addMergedRegion(cra);
        MergeRangeStyle(sheet, cra);
    }

    private static void copyStyle(Sheet sheet, int sourceRowIndex, int sourceColIndex, int targetRowIndex, int targetColIndex){
        Row sourceRow = sheet.getRow(sourceRowIndex);
        Row targetRow = sheet.getRow(targetRowIndex);

        Cell sourceCell = sourceRow.getCell(sourceColIndex);

        Cell targetCell = targetRow.getCell(targetColIndex);
        if (targetCell == null) {
            targetCell = targetRow.createCell(targetColIndex);
        }

        CellStyle cellStyle = sourceCell.getCellStyle();
        targetCell.setCellStyle(cellStyle);
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            return  cell.getNumericCellValue() + "";
        } else if (cell.getCellType() == CellType.BOOLEAN) {
            return cell.getBooleanCellValue()+"";
        } else if (cell.getCellType() == CellType.FORMULA) {
            return cell.getCellFormula();
        } else if (cell.getCellType() == CellType.ERROR) {
            return cell.getErrorCellValue()+"";
        }else if (cell.getCellType() == CellType.BLANK) {
            return "";
        }else if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else {
            return cell.getStringCellValue();
        }
    }

    private static void MergeRangeStyle(Sheet sheet, CellRangeAddress cra) {
        RegionUtil.setLeftBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
        RegionUtil.setBottomBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
        RegionUtil.setRightBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
        RegionUtil.setTopBorderColor(IndexedColors.BLACK.getIndex(), cra, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, cra, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, cra, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, cra, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, cra, sheet);
    }


}
