package io.github.millij.poi.ss.reader;

import static io.github.millij.poi.util.Beans.isInstantiableType;

import java.io.InputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.handler.RowContentsHandler;
import io.github.millij.poi.ss.handler.RowListener;


/**
 * Reader impletementation of {@link Workbook} for an OOXML .xlsx file. This implementation is
 * suitable for low memory sax parsing or similar.
 * 
 * @see XlsReader
 */
public class XlsxReader extends AbstractSpreadsheetReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsxReader.class);


    // Constructor

    public XlsxReader() {
        super();
    }


    // SpreadsheetReader Impl
    // ------------------------------------------------------------------------

    @Override
    public <T> void read(Class<T> beanClz, InputStream is, RowListener<T> listener)
            throws SpreadsheetReadException {
        readBySheetNo(beanClz, is, null, listener);
    }


    @Override
    public <T> void read(Class<T> beanClz, InputStream is, int sheetNo, RowListener<T> listener)
            throws SpreadsheetReadException {
        readBySheetNo(beanClz, is, sheetNo, listener);
    }

    private <T> void readBySheetNo(Class<T> beanClz, InputStream is, Integer sheetNo, RowListener<T> listener)
            throws SpreadsheetReadException {
        // Sanity checks
        if (!isInstantiableType(beanClz)) {
            throw new IllegalArgumentException("XlsxReader :: Invalid bean type passed !");
        }

        try (final OPCPackage opcPkg = OPCPackage.open(is)) {
            // XSSF Reader
            XSSFReader xssfReader = new XSSFReader(opcPkg);

            // Content Handler
            StylesTable styles = xssfReader.getStylesTable();
            ReadOnlySharedStringsTable ssTable = new ReadOnlySharedStringsTable(opcPkg, false);
            SheetContentsHandler sheetHandler = new RowContentsHandler<T>(beanClz, listener, 0);

            ContentHandler handler = new XSSFSheetXMLHandler(styles, ssTable, sheetHandler, true);

            // XML Reader
            XMLReader xmlParser = XMLHelper.newXMLReader();
            xmlParser.setContentHandler(handler);

            // Iterate over sheets
            XSSFReader.SheetIterator worksheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            for (int i = 0; worksheets.hasNext(); i++) {
                try (InputStream sheetInpStream = worksheets.next()) {
                    // If sheetNo is specified, skip non-specific sheets.
                    if (sheetNo != null && sheetNo != i) {
                        continue;
                    }
                    LOGGER.info("Reading the sheet {}.", i);
                    // Parse Sheet
                    xmlParser.parse(new InputSource(sheetInpStream));
                }
            }
        } catch (Exception ex) {
            String errMsg = String.format("Error reading sheet %d, to Bean %s : %s", sheetNo, beanClz, ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }

}
