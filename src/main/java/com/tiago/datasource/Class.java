package com.tiago.datasource;

import java.util.List;

public class Class {
    private String name;
    private ClassTypeEnum classType;
    private Double averageGrade;
    private List<Student> students;
    private Teacher teacher;

    public Class(
        final String name,
        final ClassTypeEnum classType,
        final Double averageGrade,
        final List<Student> students,
        final Teacher teacher
    ) {

        this.name = name;
        this.classType = classType;
        this.averageGrade = averageGrade;
        this.students = students;
        this.teacher = teacher;
    }

    public String getName() {
        return this.name;
    }

    public ClassTypeEnum getClassType() {
        return this.classType;
    }

    public Double averageGrade() {
        return this.averageGrade;
    }

    public List<Student> getStudents() {
        return this.students;
    }

    public Teacher getTeacher() {
        return this.teacher;
    }

    public String getTeacherName() {
        return this.teacher.getName();
    }
}
