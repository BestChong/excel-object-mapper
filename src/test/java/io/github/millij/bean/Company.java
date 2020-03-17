package io.github.millij.bean;

import io.github.millij.poi.ss.model.annotations.Sheet;
import io.github.millij.poi.ss.model.annotations.SheetColumn;
import lombok.*;

@Data
@AllArgsConstructor
@NoArgsConstructor
@Sheet("Companies")
public class Company {

    @SheetColumn("Company Name")
    private String name;

    @SheetColumn("# of Employees")
    private Integer noOfEmployees;

    @SheetColumn("Address")
    private String address;

}
