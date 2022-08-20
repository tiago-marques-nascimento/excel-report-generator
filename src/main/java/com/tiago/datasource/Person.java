package com.tiago.datasource;

import java.time.LocalDate;

public class Person {
    private String name;
    private LocalDate birthDate;
    private String address;
    private String mobile;
    private String email;

    public Person(
        final String name,
        final LocalDate birthDate,
        final String address,
        final String mobile,
        final String email
    ) {
        this.name = name;
        this.birthDate = birthDate;
        this.address = address;
        this.mobile = mobile;
        this.email = email;
    }

    public String getName() {
        return this.name;
    }

    public LocalDate getBirthDate() {
        return this.birthDate;
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
