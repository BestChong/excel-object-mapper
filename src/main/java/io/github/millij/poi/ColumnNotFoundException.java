package io.github.millij.poi;

import lombok.Getter;

/**
 * When the sheet column/header is empty, this exception will be thrown.
 *
 * @author Fang Gang
 */
public class ColumnNotFoundException extends RuntimeException {
    private static final long serialVersionUID = 1L;

    /**
     * The row number of columns/headers.
     */
    @Getter
    private int rowNum;

    @Getter
    private String cellColumnName;

    public ColumnNotFoundException(int rowNum) {
        super(String.format(":: Missing required column in row #%d", rowNum + 1));
        this.rowNum = rowNum;
    }

    public ColumnNotFoundException(int rowNum, String cellColumnName) {
        super(String.format(":: Missing required column(%s) in row #%d", cellColumnName, rowNum + 1));
        this.rowNum = rowNum;
        this.cellColumnName = cellColumnName;
    }
}
