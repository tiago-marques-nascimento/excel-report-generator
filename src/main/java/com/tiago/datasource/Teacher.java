package com.tiago.datasource;

import java.util.List;

public class Teacher extends Person {
    private List<Class> classes;

    public Teacher(
        final String name,
        final Integer age,
        final String address,
        final String mobile,
        final String email,
        final List<Class> classes
    ) {
        super(
            name,
            age,
            address,
            mobile,
            email
        );

        this.classes = classes;
    }

    public List<Class> getClasses() {
        return this.classes;
    }
}
