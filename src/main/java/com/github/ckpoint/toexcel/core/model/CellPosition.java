package com.github.ckpoint.toexcel.core.model;

import com.github.ckpoint.toexcel.core.type.SheetDirection;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class CellPosition {

    private int rowPosition;
    private int cellPosition;
    private SheetDirection sheetDirection;
    private final Sheet _sheet;

    public CellPosition(Sheet _sheet) {
        this(_sheet, SheetDirection.HORIZON);
    }

    public CellPosition(Sheet _sheet, SheetDirection sheetDirection) {
        this._sheet = _sheet;
        this.sheetDirection = sheetDirection;
        if (this._sheet.getRow(0) == null) {
            this._sheet.createRow(0);
        }
    }

    public void newLine() {
        if (sheetDirection.equals(SheetDirection.HORIZON)) {
            this.rowPosition++;
            this.cellPosition = 0;
        } else {
            this.cellPosition++;
            this.rowPosition = 0;
        }
    }

    public Cell nextCell() {
        Row currentRow = _sheet.getRow(rowPosition);
        if (currentRow == null) {
            return _sheet.createRow(rowPosition).createCell(cellPosition);
        } else if (currentRow.getCell(cellPosition) == null) {
            return currentRow.createCell(cellPosition);
        }

        while (true) {
            cellCountPlus();
            if (_sheet.getRow(rowPosition) == null) {
                return _sheet.createRow(rowPosition).createCell(cellPosition);
            } else if (_sheet.getRow(rowPosition).getCell(cellPosition) == null) {
                return _sheet.getRow(rowPosition).createCell(cellPosition);
            }
        }
    }

    public void clear() {
        this.rowPosition = 0;
        this.cellPosition = 0;
    }

    public SheetDirection updateDirection(SheetDirection sheetDirection) {
        this.sheetDirection = sheetDirection;
        return this.sheetDirection;
    }

    public List<Cell> skip(int cnt) {
        return IntStream.range(0, cnt).mapToObj(i -> nextCell()).collect(Collectors.toList());
    }

    public List<Cell> merge(int width, int height) {
        int targetCell = cellPosition + width - 1;
        int targetRow = rowPosition + height - 1;
        List<Cell> mergeCellList = new ArrayList<>();
        Cell originCell = null;

        for (int rowIdx = rowPosition; rowIdx < targetRow; rowIdx++) {
            Row row = _sheet.getRow(rowIdx) == null ? _sheet.createRow(rowIdx) : _sheet.getRow(rowIdx);
            for (int cellIdx = cellPosition; cellIdx < targetCell; cellIdx++) {
                if (row.getCell(cellIdx) == null) {
                    mergeCellList.add(row.createCell(cellIdx));
                } else {
                    originCell = row.getCell(cellIdx);
                    mergeCellList.add(originCell);
                }
            }
        }
        _sheet.addMergedRegion(new CellRangeAddress(rowPosition, targetRow, cellPosition, targetCell));

        if (this.sheetDirection.equals(SheetDirection.HORIZON)) {
            this.cellPosition = targetCell;
        } else {
            this.rowPosition = targetRow;
        }

        if (originCell != null) {
            for (Cell cell : mergeCellList) {
                cell.setCellValue(originCell.getStringCellValue());
                cell.setCellStyle(originCell.getCellStyle());
            }
        }
        return mergeCellList;
    }

    private void cellCountPlus() {
        if (this.sheetDirection.equals(SheetDirection.HORIZON)) {
            this.cellPosition++;
        } else {
            this.rowPosition++;
        }
    }


}
