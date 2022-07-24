package com.tiago.utils;

import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Optional;

public class ObjectAccessor {

    private Object object;

    public ObjectAccessor(final Object object) {
        this.object = object;
    }

    public String getValue() {
        return this.object.toString();
    }

    public String getValueFromDate(String format) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(format);
        if(this.object instanceof LocalDate) {
            return ((LocalDate)this.object).format(formatter);
        } else if(this.object instanceof LocalDateTime) {
            return ((LocalDateTime)this.object).format(formatter);
        } else {
            return null;
        }
    }

    public Integer getValueAsInteger() {
        return Integer.parseInt(this.object.toString());
    }

    public Double getValueAsDouble() {
        return Double.parseDouble(this.object.toString());
    }

    public Optional<ObjectAccessor> access(String methodOrField) {
        try {

            try {
                return Optional.of(new ObjectAccessor(
                    object.getClass()
                        .getMethod(methodOrField)
                        .invoke(this.object)
                ));
            } catch (
                IllegalAccessException |
                IllegalArgumentException |
                InvocationTargetException |
                NoSuchMethodException |
                SecurityException e
            ) {
                return Optional.of(new ObjectAccessor(
                    object.getClass()
                        .getMethod("get" +
                            methodOrField.substring(0, 1).toUpperCase() +
                            methodOrField.substring(1))
                        .invoke(this.object)
                ));
            }

        } catch (
            IllegalAccessException |
            IllegalArgumentException |
            InvocationTargetException |
            NoSuchMethodException |
            SecurityException e
        ) {
            e.printStackTrace();
            return Optional.empty();
        }
    }

    @SuppressWarnings("unchecked")
    public Integer listSize() {
        return ((List<Object>)object).size();
    }

    @SuppressWarnings("unchecked")
    public Optional<ObjectAccessor> accessList(Integer index) {
        return Optional.of(new ObjectAccessor(
            ((List<Object>)object).get(index)));
    }
}
