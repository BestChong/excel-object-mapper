package io.github.millij.poi.ss.handler;

import io.github.millij.poi.ss.model.ColumnMapping;
import io.github.millij.poi.ss.model.SheetRow;
import io.github.millij.poi.util.Spreadsheet;
import lombok.extern.slf4j.Slf4j;

/**
 * Row contents handler.
 *
 * @param <T> Class Type
 * @author milli, Fang Gang
 */
@Slf4j
public class RowContentsHandler<T> extends AbstractSheetContentsHandler {

    private final Class<T> beanClz;

    private final int headerRow;

    private ColumnMapping<T> columnMapping;

    private final RowListener<T> rowListener;


    // Constructors
    // ------------------------------------------------------------------------

    public RowContentsHandler(Class<T> beanClz, RowListener<T> rowListener) {
        this(beanClz, rowListener, 0);
    }

    public RowContentsHandler(Class<T> beanClz, RowListener<T> rowListener, int headerRow) {
        super();

        this.beanClz = beanClz;
        this.headerRow = headerRow;
        this.rowListener = rowListener;
    }


    // AbstractSheetContentsHandler Methods
    // ------------------------------------------------------------------------

    @Override
    void beforeRowStart(int rowNum) {
        log.debug("Start reading row - {}.", rowNum);
    }


    @Override
    void afterRowEnd(final SheetRow sheetRow) {
        // Sanity Checks
        if (sheetRow == null || sheetRow.isEmpty()) {
            return;
        }

        final int rowNum = sheetRow.getRowNum();

        // Skip rows before header row
        if (rowNum < headerRow) {
            return;
        }

        if (rowNum == headerRow) {
            columnMapping = new ColumnMapping(beanClz, sheetRow);
            return;
        }

        // Row As Bean
        T rowBean = Spreadsheet.rowAsBean(sheetRow, columnMapping);

        // Row Callback
        try {
            rowListener.row(rowNum, rowBean);
        } catch (Exception ex) {
            String errMsg = String.format("Error calling listener callback row - %d, bean - %s", rowNum, rowBean);
            throw new RuntimeException(errMsg, ex);
        }
    }

}
