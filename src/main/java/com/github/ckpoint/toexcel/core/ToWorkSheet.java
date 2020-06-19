package com.github.ckpoint.toexcel.core;

import com.github.ckpoint.toexcel.annotation.ExcelHeader;
import com.github.ckpoint.toexcel.core.converter.ExcelHeaderConverter;
import com.github.ckpoint.toexcel.core.converter.ExcelHeaderDefatulConverter;
import com.github.ckpoint.toexcel.core.model.CellPosition;
import com.github.ckpoint.toexcel.core.model.ToTitleKey;
import com.github.ckpoint.toexcel.core.style.ToWorkBookStyle;
import com.github.ckpoint.toexcel.core.type.SheetDirection;
import com.github.ckpoint.toexcel.core.type.ToWorkCellType;
import com.github.ckpoint.toexcel.exception.SheetNotFoundException;
import com.github.ckpoint.toexcel.util.ExcelHeaderHelper;
import com.github.ckpoint.toexcel.util.ModelMapperGenerator;
import com.github.ckpoint.toexcel.util.TitleRowHelper;
import lombok.Getter;
import lombok.NonNull;
import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.StringUtil;

import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * The type To work sheet.
 */
public class ToWorkSheet implements ExcelHeaderHelper, TitleRowHelper {

    private final static int DEFAULT_WIDTH_SIZE = 2560;

    @Getter
    private final ToWorkBook workBook;

    @Getter
    private final String name;
    private final Workbook _wb;
    private final CellPosition cellPosition;

    private ExcelHeaderConverter __excelHeaderConverter = new ExcelHeaderDefatulConverter();

    private Sheet _sheet;

    /**
     * Instantiates a new To work sheet.
     *
     * @param toWorkBook the to work book
     * @param wb         the wb
     */
    public ToWorkSheet(@NonNull ToWorkBook toWorkBook, @NonNull Workbook wb) {
        this(toWorkBook, wb, null);
    }

    /**
     * Instantiates a new To work sheet.
     *
     * @param toWorkBook the to work book
     * @param wb         the wb
     * @param name       the name
     */
    public ToWorkSheet(@NonNull ToWorkBook toWorkBook, @NonNull Workbook wb, String name) {
        this.workBook = toWorkBook;
        this._wb = wb;

        if (name != null && !name.isEmpty()) {
            this.name = name;
            this._sheet = this._wb.createSheet(name);
        } else {
            this._sheet = this._wb.createSheet();
            this.name = this._sheet.getSheetName();
        }
        this.cellPosition = new CellPosition(_sheet);
    }

    /**
     * Instantiates a new To work sheet.
     *
     * @param toWorkBook the to work book
     * @param _sheet     the sheet
     */
    public ToWorkSheet(@NonNull ToWorkBook toWorkBook, @NonNull Sheet _sheet) {
        this.workBook = toWorkBook;
        this._wb = _sheet.getWorkbook();
        this._sheet = _sheet;
        this.name = _sheet.getSheetName();
        this.cellPosition = new CellPosition(_sheet);
    }

    /**
     * Update header excel converter to work sheet.
     *
     * @param excelHeaderConverter the excel header converter
     * @return the to work sheet
     */
    public ToWorkSheet updateHeaderExcelConverter(ExcelHeaderConverter excelHeaderConverter) {
        this.__excelHeaderConverter = excelHeaderConverter;
        return this;
    }

    /**
     * Create title cell list.
     *
     * @param width  the width
     * @param values the values
     * @return the list
     */
    public List<ToWorkCell> createTitleCell(int width, String... values) {
        if (values == null || values.length < 1) {
            return new ArrayList<>();
        }

        List<ToWorkCell> cells = new ArrayList<>();

        for (String value : values) {
            Cell pcell = this.cellPosition.nextCell();
            _sheet.setColumnWidth(pcell.getColumnIndex(), DEFAULT_WIDTH_SIZE * width);
            cells.add(new ToWorkCell(this, pcell, value, ToWorkCellType.TITLE));
        }
        return cells;
    }

    /**
     * Create cell to newline list.
     *
     * @param values the values
     * @return the list
     */
    public List<ToWorkCell> createCellToNewline(Object... values) {
        this.newLine();
        return createCell(values);
    }

    /**
     * Create cell list.
     *
     * @param style  the style
     * @param values the values
     * @return the list
     */
    public List<ToWorkCell> createCell(@NonNull ToWorkBookStyle style, Object... values) {
        if (values == null || values.length < 1) {
            return new ArrayList<>();
        }
        return Arrays.stream(values).map(v -> new ToWorkCell(this, cellPosition.nextCell(), v, style)).collect(Collectors.toList());
    }

    /**
     * Create cell list.
     *
     * @param values the values
     * @return the list
     */
    public List<ToWorkCell> createCell(Object... values) {
        if (values == null || values.length < 1) {
            return new ArrayList<>();
        }
        return Arrays.stream(values).map(v -> new ToWorkCell(this, cellPosition.nextCell(), v)).collect(Collectors.toList());
    }

    /**
     * New line.
     */
    public void newLine() {
        this.cellPosition.newLine();
    }

    /**
     * Map list.
     *
     * @param <T>  the type parameter
     * @param type the type
     * @return the list
     */
    public <T> List<T> map(Class<T> type) {
        return map(type, null);
    }

    public <T> List<T> map(Class<T> type, Integer maxRowCnt) {
        Row titleRow = findTitleRow(type, this._sheet, this.__excelHeaderConverter);

        Map<Integer, String> excelTitleMap = new HashMap<>();
        for (int i = 0; i < titleRow.getLastCellNum(); i++) {
            excelTitleMap.put(i, titleRow.getCell(i) != null ? titleRow.getCell(i).getStringCellValue() : "");
        }

        List<String> existTitles = IntStream.range(0, titleRow.getLastCellNum()).mapToObj(titleRow::getCell)
                .filter(Objects::nonNull)
                .map(Cell::getStringCellValue).collect(Collectors.toList());

        List<Field> fields = getDeclaredFields(type);
        Set<ToTitleKey> keyset = new HashSet<>();
        for (Field field : fields) {
            if (field.getAnnotation(ExcelHeader.class) != null) {
                keyset.add(new ToTitleKey(field, existTitles, __excelHeaderConverter));
            }
        }

        List<Map<String, Object>> proxyMapList = IntStream.range(titleRow.getRowNum() + 1, getLastDataRowNum(titleRow, maxRowCnt))
                .mapToObj(this._sheet::getRow).map(row -> rowToMap(row, excelTitleMap, keyset)).collect(Collectors.toList());

        return proxyMapList.stream().filter(map -> !isEmptyValueMap(map))
                .map(map -> ModelMapperGenerator.enableFieldModelMapper().map(map, type)).collect(Collectors.toList());
    }
    private boolean isEmptyValueMap(Map<String, Object> map){
        if( map == null || map.isEmpty()){
            return true;
        }
        for (Map.Entry<String, Object> entry: map.entrySet()) {
            if( entry.getValue() == null){ continue; }
            if(!String.valueOf(entry.getValue()).trim().isEmpty()){
                return false;
            }

        }
        return true;
    }

    private int getLastDataRowNum(Row titleRow, Integer maxRowCnt) {
        if (maxRowCnt == null) {
            return this._sheet.getLastRowNum() + 1;
        }
        int firstDataRow = titleRow.getRowNum() + 1;
        return firstDataRow + maxRowCnt;
    }

    private List<Field> getDeclaredFields(Class type) {
        List<Field> fields = Arrays.stream(type.getDeclaredFields()).collect(Collectors.toList());
        if (type.getSuperclass() != null && !type.getSuperclass().equals(Object.class)) {
            fields.addAll(getDeclaredFields(type.getSuperclass()));
        }
        return fields;
    }

    /**
     * Update direction to work sheet.
     *
     * @param sheetDirection the sheet direction
     * @return the to work sheet
     */
    public ToWorkSheet updateDirection(SheetDirection sheetDirection) {
        this.cellPosition.updateDirection(sheetDirection);
        return this;
    }

    /**
     * Skip list.
     *
     * @param cnt the cnt
     * @return the list
     */
    public List<ToWorkCell> skip(int cnt) {
        List<Cell> cells = this.cellPosition.skip(cnt);
        return cells.stream().map(v -> new ToWorkCell(this, v, null)).collect(Collectors.toList());
    }

    /**
     * Merge.
     *
     * @param width  the width
     * @param height the height
     */
    public void merge(int width, int height) {
        this.cellPosition.merge(width, height);
    }

    public void writeTitle(String title, int width){
        createCell( title);
        merge(width, 1);
        newLine();
    }
    /**
     * From.
     *
     * @param list the list
     */
    public void from( List list) {
        from(null, list);
    }
    /**
     * From.
     *
     * @param list the list
     */
    public void from(String title, List list) {
        if (list == null || list.isEmpty()) {
            throw new SheetNotFoundException("Not found sheet list");
        }

        this.clear();

        Object obj = list.get(0);
        List<Field> fields = getDeclaredFields(obj.getClass());
        AtomicInteger fieldCount = new AtomicInteger();
        List<ToTitleKey> keys = fields.stream().filter(field -> field.getAnnotation(ExcelHeader.class) != null)
                .map(field -> new ToTitleKey(field, fieldCount.getAndIncrement(), __excelHeaderConverter))
                .sorted().collect(Collectors.toList());
        if( title != null){
            writeTitle(title, keys.size());
        }
        keys.forEach(key -> this.createTitleCell(1, key.getViewName()));
        for (Object o : list) {
            writeObject(o, keys);
        }
    }

    /**
     * Gets cell.
     *
     * @param rowIdx  the row idx
     * @param cellIdx the cell idx
     * @return the cell
     */
    public Cell getCell(@NonNull int rowIdx, @NonNull int cellIdx) {
        return this.cellPosition.getCell(rowIdx, cellIdx);
    }

    /**
     * Gets merged regions.
     *
     * @return the merged regions
     */
    public List<CellRangeAddress> getMergedRegions() {
        if (this._sheet.getNumMergedRegions() < 1) {
            return new ArrayList<>();
        }

        return IntStream.range(0, this._sheet.getNumMergedRegions()).mapToObj(i -> this._sheet.getMergedRegion(i)).collect(Collectors.toList());
    }

    /**
     * Gets row count.
     *
     * @return the row count
     */
    public int getRowCount() {
        if (this._sheet.getLastRowNum() == 0 && this._sheet.getRow(0) == null) {
            return 0;
        }
        return this._sheet.getLastRowNum() + 1;
    }

    private void writeObject(Object obj, List<ToTitleKey> keys) {
        this.newLine();
        keys.forEach(key -> {
            Object value = "";
            try {
                key.getField().setAccessible(true);
                value = key.getField().get(obj);
            } catch (IllegalAccessException e) {
                //TODO
            }
            this.createCell(value);
        });
    }

    private void clear() {

        for (int i = 0; i < this._sheet.getLastRowNum(); i++) {
            this._sheet.removeRow(this._sheet.getRow(i));
        }
        this.cellPosition.clear();
    }

    private Map<String, Object> rowToMap(Row row, Map<Integer, String> titleMaps, Set<ToTitleKey> setKeys) {

        Map<String, Object> valueMap = new HashMap<>();
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Object value = getCellValue(row.getCell(i));
            String title = titleMaps.get(i);
            if (title == null) {
                break;
            }
            ToTitleKey toTitleKey = setKeys.stream().filter(key -> key.isMyName(title)).findFirst().orElse(null);
            if (toTitleKey == null) {
                continue;
            }

            valueMap.put(toTitleKey.getKey(), value);
        }
        return valueMap;
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_FORMULA:
                return cell.getCachedFormulaResultType();

            case Cell.CELL_TYPE_BLANK:
            case Cell.CELL_TYPE_ERROR:
                return "";

            case Cell.CELL_TYPE_STRING:
            default:
                return cell.getStringCellValue();
        }
    }

    private Row getLastRow() {
        return this._sheet.getRow(this._sheet.getLastRowNum());
    }
}
