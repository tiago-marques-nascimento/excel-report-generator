package com.tiago.datasource;

public class Person {
    private String name;
    private Integer age;
    private String address;
    private String mobile;
    private String email;

    public Person(
        final String name,
        final Integer age,
        final String address,
        final String mobile,
        final String email
    ) {
        this.name = name;
        this.age = age;
        this.address = address;
        this.mobile = mobile;
        this.email = email;
    }

    public String getName() {
        return this.name;
    }

    public Integer getAge() {
        return this.age;
    }

    public String getAddress() {
        return this.address;
    }

    public String getMobile() {
        return this.mobile;
    }

    public String getEmail() {
        return this.email;
    }
}
