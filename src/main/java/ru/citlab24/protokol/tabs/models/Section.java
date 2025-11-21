package ru.citlab24.protokol.tabs.models;

import java.io.Serializable;

public class Section implements Serializable {
    private static final long serialVersionUID = 1L;
    private int id;         // id записи в БД (table section)
    private String name;    // название секции (подъезда)
    private int position;   // порядок показа (0,1,2,...)

    public Section() {}
    public Section(String name, int position) {
        this.name = name;
        this.position = position;
    }

    public int getId() { return id; }
    public void setId(int id) { this.id = id; }

    public String getName() { return name; }
    public void setName(String name) { this.name = name; }

    public int getPosition() { return position; }
    public void setPosition(int position) { this.position = position; }

    @Override public String toString() { return name; }
}
