package com.red.redxls.config;

import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.RowWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.handler.context.RowWriteHandlerContext;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

/**
 * @author pjh
 * @created 2024/7/23
 */
public class BlackBorderCellWriteHandler implements RowWriteHandler {

    int firstRowIdx = -1;
    int lastRowIdx = -1;
    int firstColumnIdx= -1;
    int lastColumnIdx = -1;

    public BlackBorderCellWriteHandler(int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex) {
        firstRowIdx = firstRowIndex;
        lastRowIdx = lastRowIndex;
        firstColumnIdx = firstColumnIndex;
        lastColumnIdx = lastColumnIndex;
    }

    @Override
    public void afterRowCreate(RowWriteHandlerContext context){


    }

    @Override
    public void beforeRowCreate(RowWriteHandlerContext context) {
        RowWriteHandler.super.beforeRowCreate(context);
    }

    @Override
    public void beforeRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Integer rowIndex, Integer relativeRowIndex, Boolean isHead) {
        RowWriteHandler.super.beforeRowCreate(writeSheetHolder, writeTableHolder, rowIndex, relativeRowIndex, isHead);
    }

    @Override
    public void afterRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
        RowWriteHandler.super.afterRowCreate(writeSheetHolder, writeTableHolder, row, relativeRowIndex, isHead);
    }

    @Override
    public void afterRowDispose(RowWriteHandlerContext context) {

        // This method is called before a cell is written
        Row row = context.getRow();
        short firstCellNum = row.getFirstCellNum();
        short lastCellNum = row.getLastCellNum();
        for (int i = firstCellNum; i < lastCellNum; i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                cell = row.createCell(i);
            }
            CellStyle cellStyle = cell.getCellStyle();
            if(cellStyle == null){
                Workbook workbook = context.getWriteSheetHolder().getSheet().getWorkbook();
                cellStyle = workbook.createCellStyle();
            }

            // Set the border style
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);

            // Set the black color style
            cellStyle.setTopBorderColor(IndexedColors.RED.getIndex());
            cellStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.RED.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.RED.getIndex());

            // Apply the style to the cell
            cell.setCellStyle(cellStyle);

        }

        RowWriteHandler.super.afterRowDispose(context);
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
        RowWriteHandler.super.afterRowDispose(writeSheetHolder, writeTableHolder, row, relativeRowIndex, isHead);
    }
}