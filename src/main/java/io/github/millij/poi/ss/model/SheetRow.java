package io.github.millij.poi.ss.model;

import io.github.millij.poi.ColumnDuplicateException;
import io.github.millij.poi.ColumnNotFoundException;
import io.github.millij.poi.UnsupportedException;
import io.github.millij.poi.util.Spreadsheet;
import lombok.*;
import lombok.experimental.FieldDefaults;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.util.CellAddress;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * Sheet row wrapper.
 *
 * @author Fang Gang
 */
@Slf4j
@ToString
@FieldDefaults(level = AccessLevel.PRIVATE)
public class SheetRow {

    /**
     * Row number.
     */
    @Getter
    int rowNum;

    /**
     * Cell column reference to cell object.
     * <p>
     * Example:
     * <pre>
     * "A" -> Cell("1000"), "B" -> Cell("James")
     * </pre>
     */
    Map<String, Cell> cellColumnRefToCell;

    public SheetRow(int rowNum) {
        this.rowNum = rowNum;
        this.cellColumnRefToCell = new HashMap<>();
    }

    /**
     * Put a cell object to sheet row object.
     *
     * @param cellReference cell reference, ex. A1, B3.
     * @param cellValue     cell value
     */
    public void addCell(@NonNull String cellReference, Object cellValue) {
        log.debug("Put cell: {}, {}", cellReference, cellValue);
        Cell cell = new Cell(cellReference, cellValue);
        String colRef = Spreadsheet.getCellColumnReference(cellReference);

        cellColumnRefToCell.put(colRef, cell);
    }

    /**
     * Put a cell object to sheet row object.
     *
     * @param cellAddress CellAddress object
     * @param cellValue   cell value
     */
    public void addCell(@NonNull CellAddress cellAddress, Object cellValue) {
        log.debug("Put cell: {}, {}", cellAddress, cellValue);
        Cell cell = new Cell(cellAddress, cellValue);
        String colRef = Spreadsheet.getCellColumnReference(cellAddress.toString());

        cellColumnRefToCell.put(colRef, cell);
    }

    /**
     * Get a cell object by cell column reference.
     *
     * @param cellColRef cell column reference.
     * @return return a sheet cell model.
     */
    public Cell getCell(@NonNull String cellColRef) {
        Cell cell = cellColumnRefToCell.get(cellColRef);
        return cell == null ? new Cell() : cell;
    }

    public Set<String> getCellColRefs() {
        return cellColumnRefToCell.keySet();
    }

    public int getPhysicalRowNum() {
        return rowNum + 1;
    }

    /**
     * Get column name to reference mapping, just for the header row.
     *
     * @return return column name to reference mapping
     */
    public Map<String, String> getColumnNameToReferenceMap() {
        final SheetRow headerRow = this;
        if (headerRow.isEmpty()) {
            throw new ColumnNotFoundException(headerRow.getRowNum());
        }
        Map<String, String> mapping = new HashMap<>(cellColumnRefToCell.size());
        for (String ref : headerRow.getCellColRefs()) {
            final Cell cell = headerRow.getCell(ref);
            // Cell value in header row is the column/header name
            final String columnName = String.valueOf(cell.getValue());
            // Column refuse duplication
            if (mapping.get(columnName) != null) {
                throw new ColumnDuplicateException(cell.getAddress().toString(), columnName);
            }
            mapping.put(columnName, ref);
        }
        return Collections.unmodifiableMap(mapping);
    }

    /**
     * Determine if all cells in a row are valid.
     *
     * @return Return true if there are valid cells in the row.
     */
    public boolean isEmpty() {
        return cellColumnRefToCell == null || cellColumnRefToCell.isEmpty();
    }

    public static SheetRow buildFromHSSFRow(HSSFRow hssfRow) {
        // Sanity checks
        if (hssfRow == null) {
            return null;
        }

        SheetRow sheetRow = new SheetRow(hssfRow.getRowNum());

        for (org.apache.poi.ss.usermodel.Cell cell : hssfRow) {
            // Process cell value
            switch (cell.getCellType()) {
                case STRING:
                    sheetRow.addCell(cell.getAddress(), cell.getStringCellValue());
                    break;
                case NUMERIC:
                    sheetRow.addCell(cell.getAddress(), cell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    sheetRow.addCell(cell.getAddress(), cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    var err1 = String.format("Formula cell(%s) type not support.", cell.getAddress());
                    throw new UnsupportedException(err1);
                case BLANK:
                    log.warn("Cell(%s) data is BLANK.", cell.getAddress());
                    break;
                case ERROR:
                    var err2 = String.format("Cell(%s) data not support.", cell.getAddress());
                    throw new UnsupportedException(err2);
                default:
                    break;
            }
        }
        return sheetRow;
    }

    /**
     * Sheet cell model.
     */
    @Getter
    @Setter
    @ToString
    @AllArgsConstructor
    @NoArgsConstructor
    @FieldDefaults(level = AccessLevel.PRIVATE)
    public static class Cell {

        /**
         * Cell Address
         */
        CellAddress address;

        /**
         * Cell formatted value object
         */
        Object value;

        public Cell(String cellReference, Object value) {
            if (StringUtils.isBlank(cellReference)) {
                throw new IllegalArgumentException("For input string: " + cellReference);
            }
            try {
                this.address = new CellAddress(cellReference);
            } catch (NumberFormatException e) {
                String errMag = String.format("Cell address format failed. cell reference: %s", cellReference);
                throw new UnsupportedException(errMag);
            }
            this.value = value;
        }
    }
}
