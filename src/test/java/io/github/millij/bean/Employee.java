package io.github.millij.bean;

import io.github.millij.poi.ss.model.annotations.Sheet;
import io.github.millij.poi.ss.model.annotations.SheetColumn;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString
@Sheet
public class Employee {

    // Note that Id and Name are annotated at name level
    @SheetColumn(value = "ID", nullable = false)
    private String id;
    @SheetColumn(value = "Name")
    private String name;

    @SheetColumn("Age")
    private Integer age;

    @SheetColumn("Gender")
    private String gender;

    @SheetColumn("Height (mts)")
    private Double height;

    @SheetColumn("Address")
    private String address;


    // Constructors
    // ------------------------------------------------------------------------

    public Employee() {
        // Default
    }

    public Employee(String id, String name, Integer age, String gender, Double height) {
        super();

        this.id = id;
        this.name = name;
        this.age = age;
        this.gender = gender;
        this.height = height;
    }

}
