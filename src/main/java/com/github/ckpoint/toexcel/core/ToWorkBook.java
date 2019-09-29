package com.github.ckpoint.toexcel.core;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class ToWorkBook {

    private final HSSFWorkbook _wb;

    private final Map<ToWorkBookStyle, CellStyle> _styleMap = new HashMap<>();

    public static final Integer MAX_EXCEL_ROW = 65535;

    private List<ToWorkSheet> sheets = new ArrayList<>();

    public ToWorkBook() {
        this._wb = new HSSFWorkbook();
    }

    public ToWorkBook(File file) throws IOException {
        if (!file.exists() || file.isDirectory()) {
            throw new FileNotFoundException();
        }

        this._wb = new HSSFWorkbook(new FileInputStream(file));
        this.__initSheet();
    }

    public ToWorkBook(FileInputStream fis) throws IOException {
        this._wb = new HSSFWorkbook(fis);
        this.__initSheet();
    }

    public ToWorkSheet createSheet() {
        ToWorkSheet sheet = new ToWorkSheet(this, this._wb.createSheet());
        this.sheets.add(sheet);
        return sheet;
    }

    public ToWorkSheet createSheet(String name) {
        ToWorkSheet sheet = new ToWorkSheet(this, this._wb.createSheet(name));
        this.sheets.add(sheet);
        return sheet;
    }


    public void writeFile(String filePath) throws IOException {
        File file = new File(filePath);
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        this._wb.write(fileOutputStream);
    }

    public int sheetSize() {
        return this.sheets.size();
    }

    public ToWorkSheet getSheetAt(int idx) {
        return this.sheets.get(idx);
    }

    public ToWorkSheet getSheet(String name) {
        return this.sheets.stream().filter(st -> name.equalsIgnoreCase(st.getName())).findFirst().orElse(null);
    }

    public synchronized CellStyle createStyle(ToWorkCell cell) {
        CellStyle cellStyle = _styleMap.get(cell.getStyle());
        if (cellStyle != null) {
            return cellStyle;
        }

        cellStyle = cell.getStyle().convertHssfStyle(this._wb);

        this._styleMap.put(cell.getStyle(), cellStyle);
        return cellStyle;
    }

    private void __initSheet() {
        this.sheets = IntStream.range(0, this._wb.getNumberOfSheets()).mapToObj(this._wb::getSheetAt).map(sheet -> new ToWorkSheet(this, sheet)).collect(Collectors.toList());
    }


}
