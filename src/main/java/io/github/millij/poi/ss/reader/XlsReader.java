package io.github.millij.poi.ss.reader;

import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.handler.RowListener;
import io.github.millij.poi.ss.model.ColumnMapping;
import io.github.millij.poi.ss.model.SheetRow;
import io.github.millij.poi.util.Spreadsheet;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.InputStream;

import static io.github.millij.poi.util.Beans.isInstantiableType;

/**
 * Reader impletementation of {@link Workbook} for an POIFS file (.xls).
 *
 * @author milli, Fang Gang
 * @see XlsxReader
 */
@Slf4j
public class XlsReader extends AbstractSpreadsheetReader {

    // Constructor

    public XlsReader() {
        super();
    }


    // WorkbookReader Impl
    // ------------------------------------------------------------------------

    @Override
    public <T> void read(Class<T> beanClz, InputStream is, RowListener<T> listener) throws SpreadsheetReadException {
        // Sanity checks
        if (!isInstantiableType(beanClz)) {
            throw new IllegalArgumentException("XlsReader :: Invalid bean type passed!");
        }

        String sheetName = "";

        try {
            final HSSFWorkbook wb = new HSSFWorkbook(is);
            final int sheetCount = wb.getNumberOfSheets();
            log.debug("Total no. of sheets found in HSSFWorkbook : #{}", sheetCount);

            // Iterate over sheets
            for (int i = 0; i < sheetCount; i++) {
                final HSSFSheet sheet = wb.getSheetAt(i);
                sheetName = sheet.getSheetName();
                log.debug("Processing HSSFSheet at No. : {}", i);

                // Process Sheet
                this.processSheet(beanClz, sheet, 0, listener);
            }

            // Close workbook
            wb.close();
        } catch (Exception ex) {
            log.error("HSSFSheet to Bean({}) Error - Sheet[{}] {}", beanClz.getSimpleName(), sheetName, ex.getMessage());
            throw new SpreadsheetReadException(sheetName, ex);
        }

    }

    @Override
    public <T> void read(Class<T> beanClz, InputStream is, int sheetNo, RowListener<T> listener)
            throws SpreadsheetReadException {
        // Sanity checks
        if (!isInstantiableType(beanClz)) {
            throw new IllegalArgumentException("XlsReader :: Invalid bean type passed!");
        }

        String sheetName = "";

        try {
            HSSFWorkbook wb = new HSSFWorkbook(is);
            final HSSFSheet sheet = wb.getSheetAt(sheetNo);
            sheetName = sheet.getSheetName();

            // Process Sheet
            this.processSheet(beanClz, sheet, 0, listener);

            // Close workbook
            wb.close();
        } catch (Exception ex) {
            log.error("HSSFSheet to Bean({}) Error - Sheet[{}] {}", beanClz.getSimpleName(), sheetName, ex.getMessage());
            throw new SpreadsheetReadException(sheetName, ex);
        }
    }


    // Sheet Process

    protected <T> void processSheet(Class<T> beanClz, HSSFSheet sheet, int headerRowNo, RowListener<T> rowListener) {
        // Get header row data
        final SheetRow headerRow = SheetRow.buildFromHSSFRow(sheet.getRow(headerRowNo));
        final ColumnMapping<T> columnMapping = new ColumnMapping(beanClz, headerRow);

        for (Row row : sheet) {
            // Process Row Data
            int rowNum = row.getRowNum();
            // Skip Header row
            if (rowNum <= 0) {
                continue;
            }

            SheetRow sheetRow = SheetRow.buildFromHSSFRow((HSSFRow) row);
            if (sheetRow.isEmpty()) {
                log.warn("Row(idx{}) data is empty.", row.getRowNum());
                continue;
            }

            // Row data as Bean
            T rowBean = Spreadsheet.rowAsBean(sheetRow, columnMapping);
            // Row Callback
            rowListener.row(rowNum, rowBean);
        }
    }


    // Private Methods
    // ------------------------------------------------------------------------

}
