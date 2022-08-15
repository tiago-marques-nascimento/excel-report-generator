package com.tiago.datasource;

import java.util.List;

public class DataSource {
    private String title;
    private List<Student> students;

    public DataSource(
        final String title,
        final List<Student> students
    ) {
        this.title = title;
        this.students = students;
    }

    public String getTitle() {
        return this.title;
    }

    public List<Student> getStudents() {
        return this.students;
    }
}