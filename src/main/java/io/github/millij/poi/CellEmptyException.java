package io.github.millij.poi;

/**
 * When a cell's column is marked as non-empty by <code>@SheetColumn</code>, the cell is empty,
 * and then this exception will be thrown.<br>
 * <p>
 * See {@link io.github.millij.poi.ss.model.annotations.SheetColumn}
 *
 * @author Fang Gang
 */
public class CellEmptyException extends CellException {
    private static final long serialVersionUID = 1L;

    public CellEmptyException(String cellReference, String cellColumnName) {
        super(cellReference, cellColumnName,
                String.format("Cell[%s] :: The '%s' not allow empty.", cellReference, cellColumnName));
    }
}
