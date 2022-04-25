package com.tiago.datasource;

import java.util.List;

public class Programmer {
    private String name;
    private Integer age;
    private Integer yearsExperience;
    private List<ProgrammingLanguage> programmingLanguages;

    public void setName(final String name) {
        this.name = name;
    }

    public void setAge(final Integer age) {
        this.age = age;
    }

    public void setYearsExperience(final Integer yearsExperience) {
        this.yearsExperience = yearsExperience;
    }

    public void setProgrammingLanguages(final List<ProgrammingLanguage> programmingLanguages) {
        this.programmingLanguages = programmingLanguages;
    }

    public String getName() {
        return this.name;
    }

    public Integer getAge() {
        return this.age;
    }

    public Integer getYearsExperience() {
        return this.yearsExperience;
    }

    public List<ProgrammingLanguage> getProgrammingLanguages() {
        return this.programmingLanguages;
    }

    public Programmer() {
    }
}