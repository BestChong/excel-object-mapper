package io.github.millij.poi.ss.model;

import io.github.millij.poi.ColumnNotFoundException;
import io.github.millij.poi.ss.model.annotations.SheetColumn;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.NonNull;
import lombok.ToString;
import lombok.experimental.FieldDefaults;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.util.*;

/**
 *
 * @param <T> Class Type
 * @author Fang Gang
 */
@ToString
@FieldDefaults(level = AccessLevel.PRIVATE)
public class ColumnMapping<T> {

    @Getter
    final Class<T> beanClz;
    final Map<String, Property> columnNameToProperty;

    public ColumnMapping(Class<T> beanClz, SheetRow headerRow) {
        this.beanClz = beanClz;
        this.columnNameToProperty = new HashMap<>();

        init(beanClz);
        checkAndMapping(headerRow);
    }

    /**
     * Get Bean Property by sheet column name.
     *
     * @param columnName sheet column name
     * @return return Property object.
     */
    public Property get(@NonNull String columnName) {
        return columnNameToProperty.get(columnName);
    }

    public Set<String> getCellColNames() {
        if (columnNameToProperty == null) {
            return new HashSet<>();
        }
        return columnNameToProperty.keySet();
    }

    public boolean isEmpty() {
        return columnNameToProperty == null || columnNameToProperty.isEmpty();
    }

    // Private Methods
    // ------------------------------------------------------------------------

    private void init(Class<T> beanClz) {
        // Sanity checks
        if (beanClz == null) {
            throw new IllegalArgumentException("Error :: Invalid Excel Bean Type - null");
        }

        // Fields
        Field[] fields = beanClz.getDeclaredFields();
        for (Field f : fields) {
            final SheetColumn fa = f.getAnnotation(SheetColumn.class);
            this.set(new Property(f.getName(), fa));
        }

        // Methods
        /*
        Method[] methods = beanClz.getDeclaredMethods();
        for (Method m : methods) {
            final String fieldName = Beans.getFieldName(m);
            final Boolean canCover = fieldAnnotationMap.get(fieldName);
            if (canCover == null || canCover) {
                columnNameToProperty.remove(fieldName);
                this.set(new Property(fieldName, m.getAnnotation(SheetColumn.class)));
            }
        }
        */
    }

    private void set(final Property property) {
        columnNameToProperty.put(property.getColumnName(), property);
    }

    /**
     * Check the validation of column in the header row,
     * then setup the property - columnReference.
     *
     * @param headerRow sheet header row data
     */
    private void checkAndMapping(@NonNull SheetRow headerRow) {
        // Get sheet header name to column reference mapping
        Map<String, String> nameToReferenceMap = headerRow.getColumnNameToReferenceMap();

        removeNotNeedKV(nameToReferenceMap);

        Iterator<Map.Entry<String, Property>> it = columnNameToProperty.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, Property> entry = it.next();
            final String colName = entry.getKey();
            final Property property = entry.getValue();
            final String colRef = nameToReferenceMap.get(colName);

            if (StringUtils.isNotEmpty(colRef)) {

                property.columnReference = colRef;
            } else {

                if (property.isNullable()) {
                    // If not found the reference and property.nullable=true, remove it
                    it.remove();
                    continue;
                } else {
                    throw new ColumnNotFoundException(headerRow.getRowNum(), colName);
                }
            }
        }
    }

    private void removeNotNeedKV(Map<String, String> nameToReferenceMap) {
        Map<String, String> notFound = new HashMap<>();
        Map<String, String> found = new HashMap<>();
        boolean flag = false;

        for (String colName : columnNameToProperty.keySet()) {
            final Property property = columnNameToProperty.get(colName);
            final String colRef = nameToReferenceMap.get(colName);
            if (property.exclusive) {
                flag = true;
                if (StringUtils.isNotEmpty(colRef)) {
                    found.put(colName, colRef);
                } else {
                    notFound.put(colName, colRef);
                }
            }
        }

        if (flag) {
            if (found.size() <= 0) {
                throwException(notFound, "One of the columns must be included: ");
            } else if (found.size() == 1) {
                removeExclusiveColumn(notFound);
            } else {
                throwException(found, "These columns cannot coexist: ");
            }
        }
    }

    private void throwException(Map<String, String> found, String s) {
        StringBuilder msgBuilder = new StringBuilder(s);
        final Iterator<String> iterator = found.keySet().iterator();
        while (iterator.hasNext()) {
            msgBuilder.append(iterator.next());
            if (iterator.hasNext()) {
                msgBuilder.append(", ");
            } else {
                msgBuilder.append(".");
            }
        }
        // TODO refactor it
        throw new RuntimeException(msgBuilder.toString());
    }

    private void removeExclusiveColumn(Map<String, String> notFound) {
        for (String colName : notFound.keySet()) {
            columnNameToProperty.remove(colName);
        }
    }

    private boolean columnIsNotExist(String fieldName) {
        return !columnNameToProperty.containsKey(fieldName);
    }


    @Getter
    @ToString
    @FieldDefaults(level = AccessLevel.PRIVATE)
    public static class Property {
        /**
         * Property name
         */
        String fieldName;

        /**
         * Name of the column to map the annotated property with.
         */
        String columnName;

        /**
         * The reference of the column, ex. A or F
         */
        String columnReference;

        /**
         * Property can be null
         */
        boolean nullable;

        /**
         * Exclusive column flag, columns that are also set to true cannot exist at the same time.
         */
        boolean exclusive;

        protected Property(String fieldName, SheetColumn fa) {
            if (fa == null || StringUtils.isEmpty(fa.value())) {
                this.fieldName = fieldName;
                this.columnName = fieldName;
                this.nullable = true;
                this.exclusive = false;
            } else {
                this.fieldName = fieldName;
                this.columnName = fa.value();
                this.nullable = fa.nullable();
                this.exclusive = fa.exclusive();
            }
        }
    }
}
