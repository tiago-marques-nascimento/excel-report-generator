package com.tiago.datasource;

import java.time.LocalDate;
import java.util.List;

public class Student extends Person {
    private List<Class> classes;

    public Student(
        final String name,
        final LocalDate birthDate,
        final String address,
        final String mobile,
        final String email,
        final List<Class> classes
    ) {
        super(
            name,
            birthDate,
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