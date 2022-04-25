package com.tiago.datasource;

import java.util.List;

public class DataSource {
    private List<Programmer> programmers;

    public void setProgrammers(final List<Programmer> programmers) {
        this.programmers = programmers;
    }

    public List<Programmer> getProgrammers() {
        return this.programmers;
    }

    public DataSource() {
    }
}