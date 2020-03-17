package io.github.millij.poi;


import lombok.Getter;

/**
 * Spread sheet read exception.
 *
 * @author Fang Gang
 */
public class SpreadsheetReadException extends Exception {

    private static final long serialVersionUID = 1L;

    @Getter
    private String sheetName;

    // Constructors
    // ------------------------------------------------------------------------

    public SpreadsheetReadException(String sheetName, Throwable cause) {
        super(formatMessage(sheetName, cause), cause);
        this.sheetName = sheetName;
    }

    private static String formatMessage(String sheetName, Throwable cause) {
        if (sheetName == null) {
            sheetName = "?";
        }
        return "Sheet[" + sheetName + "] " + cause.getMessage();
    }
}
