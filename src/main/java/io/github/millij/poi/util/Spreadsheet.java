package io.github.millij.poi.util;

import io.github.millij.poi.CellEmptyException;
import io.github.millij.poi.UnsupportedException;
import io.github.millij.poi.ss.model.ColumnMapping;
import io.github.millij.poi.ss.model.SheetRow;
import io.github.millij.poi.ss.model.annotations.SheetColumn;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

/**
 * Spreadsheet related utilites.
 */
@Slf4j
public final class Spreadsheet {

    private Spreadsheet() {
        // Utility Class
    }


    // Utilities
    // ------------------------------------------------------------------------

    /**
     * Splits the CellReference and returns only the column reference.
     *
     * @param cellRef the cell reference value (ex. D3)
     * @return returns the column index "D" from the cell reference "D3"
     */
    public static String getCellColumnReference(String cellRef) {
        String cellColRef = cellRef.split("[0-9]*$")[0];
        return cellColRef;
    }


    // Bean :: Property Utils

    public static Map<String, String> getPropertyToColumnNameMap(Class<?> beanType) {
        // Sanity checks
        if (beanType == null) {
            throw new IllegalArgumentException("getPropertyToColumnNameMap :: Invalid ExcelBean type - " + beanType);
        }

        // Property to Column name Mapping
        final Map<String, String> mapping = new HashMap<String, String>();

        // Fields
        Field[] fields = beanType.getDeclaredFields();
        for (Field f : fields) {
            String fieldName = f.getName();
            mapping.put(fieldName, fieldName);

            SheetColumn ec = f.getAnnotation(SheetColumn.class);
            if (ec != null && StringUtils.isNotEmpty(ec.value())) {
                mapping.put(fieldName, ec.value());
            }
        }

        // Methods
        Method[] methods = beanType.getDeclaredMethods();
        for (Method m : methods) {
            String fieldName = Beans.getFieldName(m);
            if (!mapping.containsKey(fieldName)) {
                mapping.put(fieldName, fieldName);
            }

            SheetColumn ec = m.getAnnotation(SheetColumn.class);
            if (ec != null && StringUtils.isNotEmpty(ec.value())) {
                mapping.put(fieldName, ec.value());
            }
        }

        log.info("Bean property to Excel Column of - {} : {}", beanType, mapping);
        return Collections.unmodifiableMap(mapping);
    }

    public static List<String> getColumnNames(Class<?> beanType) {
        // Bean Property to Column Mapping
        final Map<String, String> propToColumnMap = getPropertyToColumnNameMap(beanType);

        final ArrayList<String> columnNames = new ArrayList<>(propToColumnMap.values());
        return columnNames;
    }

    // Read from Bean : as Row Data
    // ------------------------------------------------------------------------

    public static Map<String, String> asRowDataMap(Object beanObj, List<String> colHeaders) throws Exception {
        // Excel Bean Type
        final Class<?> beanType = beanObj.getClass();

        // RowData map
        final Map<String, String> rowDataMap = new HashMap<String, String>();

        // Fields
        for (Field f : beanType.getDeclaredFields()) {
            if (!f.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }

            String fieldName = f.getName();

            SheetColumn ec = f.getAnnotation(SheetColumn.class);
            String header = StringUtils.isEmpty(ec.value()) ? fieldName : ec.value();
            if (!colHeaders.contains(header)) {
                continue;
            }

            rowDataMap.put(header, Beans.getFieldValueAsString(beanObj, fieldName));
        }

        // Methods
        for (Method m : beanType.getDeclaredMethods()) {
            if (!m.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }

            String fieldName = Beans.getFieldName(m);

            SheetColumn ec = m.getAnnotation(SheetColumn.class);
            String header = StringUtils.isEmpty(ec.value()) ? fieldName : ec.value();
            if (!colHeaders.contains(header)) {
                continue;
            }

            rowDataMap.put(header, Beans.getFieldValueAsString(beanObj, fieldName));
        }

        return rowDataMap;
    }


    // Write to Bean :: from Row data
    // ------------------------------------------------------------------------

    public static <T> T rowAsBean(SheetRow sheetRow, ColumnMapping<T> columnMapping) {

        final Class<T> beanClz = columnMapping.getBeanClz();
        // Sanity checks
        if (beanClz == null || columnMapping.isEmpty() || sheetRow.isEmpty()) {
            log.warn("Mapping row as bean failed, beanClz - {}, colRefToProperty - {}, dataRow - {}.",
                    beanClz.getSimpleName(), columnMapping, sheetRow);
            return null;
        }

        T rowBean = newBeanInstance(beanClz);

        // Fill in the data
        for (String colName : columnMapping.getCellColNames()) {

            ColumnMapping.Property property = columnMapping.get(colName);
            final String colRef = property.getColumnReference();
            Object cellValue = sheetRow.getCell(colRef).getValue();

            if (!property.isNullable()) {
                if (cellValue == null || StringUtils.isEmpty(cellValue.toString())) {
                    final String cellRef = colRef + sheetRow.getPhysicalRowNum();
                    throw new CellEmptyException(cellRef, property.getColumnName());
                }
            }

            try {
                // Set the property value in the current row object bean
                BeanUtils.setProperty(rowBean, property.getFieldName(), cellValue);
            } catch (IllegalAccessException | InvocationTargetException ex) {
                String errMsg = String.format("Failed to set bean property - %s, value - %s, sheetRow - %s.",
                        property.getFieldName(), cellValue, sheetRow);
                log.error(errMsg, ex);
                throw new UnsupportedException(errMsg);
            }

        }
        return rowBean;
    }

    private static <T> T newBeanInstance(Class<T> beanClz) {
        T rowBean;
        try {
            // Create new Instance
            rowBean = beanClz.newInstance();
        } catch (InstantiationException | IllegalAccessException ex) {
            String errMsg = String.format("Error while creating bean - %s.", beanClz);
            log.error(errMsg, ex);
            throw new UnsupportedException(errMsg);
        }
        return rowBean;
    }

}
