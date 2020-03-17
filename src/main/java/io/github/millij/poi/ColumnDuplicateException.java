package io.github.millij.poi;

/**
 * When the sheet column/header is duplicate, this exception will be thrown.
 *
 * @author Fang Gang
 */
public class ColumnDuplicateException extends CellException {
    private static final long serialVersionUID = 1L;

    public ColumnDuplicateException(String cellReference, String cellColumnName) {
        super(cellReference, cellColumnName,
                String.format("Cell[%s] :: refuse duplicate header - %s.", cellReference, cellColumnName));
    }
}
