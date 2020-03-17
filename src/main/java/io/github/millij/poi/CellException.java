package io.github.millij.poi;

import lombok.Getter;


/**
 * Base cell exception.
 *
 * @author Fang Gang
 */
public class CellException extends RuntimeException {
    private static final long serialVersionUID = 1L;

    /**
     * The reference of cell.
     */
    @Getter
    private String cellReference;

    /**
     * The header/column name of cell.
     */
    @Getter
    private String cellColumnName;

    public CellException(String cellReference, String cellColumnName, String message) {
        super(message);
        this.cellReference = cellReference;
        this.cellColumnName = cellColumnName;
    }

    public CellException(String cellReference, String cellColumnName) {
        super(String.format("Cell[%s] :: error, column: %s.", cellReference, cellColumnName));
        this.cellReference = cellReference;
        this.cellColumnName = cellColumnName;
    }
}
