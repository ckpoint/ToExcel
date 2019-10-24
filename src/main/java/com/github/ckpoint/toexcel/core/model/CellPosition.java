package com.github.ckpoint.toexcel.core.model;

import com.github.ckpoint.toexcel.core.type.SheetDirection;
import com.github.ckpoint.toexcel.exception.CellNotFoundException;
import com.github.ckpoint.toexcel.exception.RowNotFoundException;
import lombok.NonNull;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * The type Cell position.
 */
public class CellPosition {

    private final Sheet _sheet;
    private int rowPosition;
    private int cellPosition;
    private SheetDirection sheetDirection;

    /**
     * Instantiates a new Cell position.
     *
     * @param _sheet the sheet
     */
    public CellPosition(@NonNull Sheet _sheet) {
        this(_sheet, SheetDirection.HORIZON);
    }

    /**
     * Instantiates a new Cell position.
     *
     * @param _sheet         the sheet
     * @param sheetDirection the sheet direction
     */
    public CellPosition(@NonNull Sheet _sheet, SheetDirection sheetDirection) {
        this._sheet = _sheet;
        this.sheetDirection = sheetDirection;
        if (this._sheet.getRow(0) == null) {
            this._sheet.createRow(0);
        }
    }

    /**
     * New line.
     */
    public void newLine() {
        if (sheetDirection.equals(SheetDirection.HORIZON)) {
            this.rowPosition++;
            this.cellPosition = 0;
        } else {
            this.cellPosition++;
            this.rowPosition = 0;
        }
    }

    /**
     * Next cell cell.
     *
     * @return the cell
     */
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

    /**
     * Clear.
     */
    public void clear() {
        this.rowPosition = 0;
        this.cellPosition = 0;
    }

    /**
     * Update direction sheet direction.
     *
     * @param sheetDirection the sheet direction
     * @return the sheet direction
     */
    public SheetDirection updateDirection(SheetDirection sheetDirection) {
        this.sheetDirection = sheetDirection;
        return this.sheetDirection;
    }

    /**
     * Skip list.
     *
     * @param cnt the cnt
     * @return the list
     */
    public List<Cell> skip(int cnt) {
        return IntStream.range(0, cnt).mapToObj(i -> nextCell()).collect(Collectors.toList());
    }

    /**
     * Merge list.
     *
     * @param width  the width
     * @param height the height
     * @return the list
     */
    public List<Cell> merge(int width, int height) {
        int targetCell = cellPosition + width - 1;
        int targetRow = rowPosition + height - 1;
        List<Cell> mergeCellList = new ArrayList<>();
        Cell originCell = this._sheet.getRow(rowPosition) == null
                ? null : this._sheet.getRow(rowPosition).getCell(cellPosition);

        for (int rowIdx = rowPosition; rowIdx <= targetRow; rowIdx++) {
            Row row = _sheet.getRow(rowIdx) == null ? _sheet.createRow(rowIdx) : _sheet.getRow(rowIdx);
            for (int cellIdx = cellPosition; cellIdx <= targetCell; cellIdx++) {
                if (row.getCell(cellIdx) == null) {
                    mergeCellList.add(row.createCell(cellIdx));
                } else {
                    originCell = originCell != null ? originCell : row.getCell(cellIdx);
                    mergeCellList.add(originCell);
                }
            }
        }
        _sheet.addMergedRegion(new CellRangeAddress(rowPosition, targetRow, cellPosition, targetCell));

        if (SheetDirection.HORIZON.equals(this.sheetDirection)) {
            this.cellPosition = targetCell;
        } else {
            this.rowPosition = targetRow;
        }

        if (originCell != null) {
            for (Cell cell : mergeCellList) {
                _sheet.setColumnWidth(cell.getColumnIndex(), _sheet.getColumnWidth(originCell.getColumnIndex()));
                cell.setCellValue(originCell.getStringCellValue());
                cell.setCellStyle(originCell.getCellStyle());
            }
        }
        return mergeCellList;
    }

    public Cell getCell(@NonNull int rowIdx, @NonNull int cellIdx) {
        Row row = this._sheet.getRow(rowIdx);
        if (row == null) {
            throw new RowNotFoundException("Not found row index " + rowIdx);
        }
        Cell cell = row.getCell(cellIdx);
        if (cell == null) {
            throw new CellNotFoundException("Not found cell index " + cellIdx);
        }
        return cell;
    }

    private void cellCountPlus() {
        if (SheetDirection.HORIZON.equals(this.sheetDirection)) {
            this.cellPosition++;
        } else {
            this.rowPosition++;
        }
    }


}
