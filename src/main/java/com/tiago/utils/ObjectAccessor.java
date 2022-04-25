package com.tiago.utils;

import java.lang.reflect.InvocationTargetException;
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

    public Optional<ObjectAccessor> access(String field) {
        try {
            return Optional.of(new ObjectAccessor(
                object.getClass()
                    .getMethod("get" +
                        field.substring(0, 1).toUpperCase() +
                        field.substring(1))
                    .invoke(this.object)));
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
