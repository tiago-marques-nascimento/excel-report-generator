package com.tiago.datasource;

public enum ClassTypeEnum {

    DAY_CLASS,
    NIGHT_CLASS;

    public String toString(ClassTypeEnum classType) {
        switch(classType) {
            case DAY_CLASS:
                return "Day class";
            case NIGHT_CLASS:
                return "Night class";
            default:
                return "";
        }
    }
}
