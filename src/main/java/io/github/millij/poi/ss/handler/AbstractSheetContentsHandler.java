package io.github.millij.poi.ss.handler;

import io.github.millij.poi.ss.model.SheetRow;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * @author milli, Fang Gang
 */
@Slf4j
abstract class AbstractSheetContentsHandler implements SheetContentsHandler {

    private SheetRow sheetRow;

    // Methods
    // ------------------------------------------------------------------------

    // Abstract

    abstract void beforeRowStart(int rowNum);

    abstract void afterRowEnd(final SheetRow sheetRow);



    // SheetContentsHandler Implementations
    // ------------------------------------------------------------------------

    @Override
    public void startRow(int rowNum) {
        // Callback
        this.beforeRowStart(rowNum);

        // Start handle row
        this.sheetRow = new SheetRow(rowNum);
    }

    @Override
    public void endRow(int rowNum) {
        // Callback
        this.afterRowEnd(sheetRow);
    }

    @Override
    public void cell(String cellRef, String cellVal, XSSFComment comment) {
        // Sanity Checks
        if (StringUtils.isEmpty(cellRef)) {
            log.error("Row[#] {} : Cell reference is empty - {}", sheetRow.getRowNum(), cellRef);
            return;
        }

        if (StringUtils.isEmpty(cellVal)) {
            log.warn("Row[#] {} - Cell[ref] formatted value is empty : {} - {}",
                    sheetRow.getRowNum(), cellRef, cellVal);
            return;
        }

        // Set the CellValue into the SheetRow
        sheetRow.addCell(cellRef, cellVal);
        log.debug("cell - Saving Column value : {} - {}", cellRef, cellVal);
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        // TODO Auto-generated method stub

    }

}
