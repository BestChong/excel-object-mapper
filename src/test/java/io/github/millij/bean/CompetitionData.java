package io.github.millij.bean;

import io.github.millij.poi.ss.model.annotations.Sheet;
import io.github.millij.poi.ss.model.annotations.SheetColumn;
import lombok.Data;

import java.util.Date;

/**
 * @author Fang Gang
 * @date 2019/10/17
 **/
@Data
@Sheet
public class CompetitionData {

    private Long id;

    private Long competitionId;

    @SheetColumn(value = "赛项")
    private String competitionType;

    @SheetColumn(value = "学员编号", exclusive = true)
    private String studentNo;

    @SheetColumn(value = "用户ID", nullable = false, exclusive = true)
    private String userId;

    @SheetColumn(value = "学员姓名", nullable = false)
    private String studentName;

    @SheetColumn(value = "总分", nullable = false)
    private Double score;

    @SheetColumn(value = "奖项", nullable = false)
    private String award;

    @SheetColumn(value = "城市", nullable = false)
    private String city;

    @SheetColumn(value = "证书编号", nullable = false)
    private String certificateNo;

    @SheetColumn
    private Date createTime;

    private String createUser;

    private Date modifyTime;

    private String modifyUser;

}
