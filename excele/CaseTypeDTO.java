package com.spr.hcase.excele;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;

import java.util.Arrays;
import java.util.Collection;
import java.util.List;

@Data
public class CaseTypeDTO {

    @ExcelProperty("委员会名称")
    private String orgName;

    @ExcelProperty({"建设工程", "数量"})
    private String constructionNum;

    @ExcelProperty({"建设工程", "标的"})
    private String constructionTarget;

    @ExcelProperty({"金融", "数量"})
    private String financeNum;

    @ExcelProperty({"金融", "标的"})
    private String financeTarget;

    @ExcelProperty({"房地产", "数量"})
    private String estateNum;

    @ExcelProperty({"房地产", "标的"})
    private String estateTarget;

    @ExcelProperty({"买卖", "数量"})
    private String saleNum;

    @ExcelProperty({"买卖", "标的"})
    private String saleTarget;

    @ExcelProperty({"租赁", "数量"})
    private String leaseNum;

    @ExcelProperty({"租赁", "标的"})
    private String leaseTarget;

    @ExcelProperty({"股权转让", "数量"})
    private String equityNum;

    @ExcelProperty({"股权转让", "标的"})
    private String equityTarget;

    @ExcelProperty({"土地交易", "数量"})
    private String landNum;

    @ExcelProperty({"土地交易", "标的"})
    private String landTarget;

    @ExcelProperty({"保险", "数量"})
    private String insuranceNum;

    @ExcelProperty({"保险", "标的"})
    private String insuranceTarget;

    @ExcelProperty({"物业", "数量"})
    private String propertyNum;

    @ExcelProperty({"物业", "标的"})
    private String propertyTarget;

    @ExcelProperty({"农业生产经营", "数量"})
    private String agricultureNum;

    @ExcelProperty({"农业生产经营", "标的"})
    private String agricultureTarget;

    @ExcelProperty({"电子商务", "数量"})
    private String ecommerceNum;

    @ExcelProperty({"电子商务", "标的"})
    private String ecommerceTarget;


    @ExcelProperty({"交通事故赔偿", "数量"})
    private String trafficAccidentsNum;

    @ExcelProperty({"交通事故赔偿", "标的"})
    private String trafficAccidentsTarget;

    @ExcelProperty({"医院纠纷", "数量"})
    private String hospitalDisputesNum;

    @ExcelProperty({"医院纠纷", "标的"})
    private String hospitalDisputesTarget;


    @ExcelProperty({"知识产权", "数量"})
    private String iprNum;

    @ExcelProperty({"知识产权", "标的"})
    private String iprTarget;

    @ExcelProperty({"其他", "数量"})
    private String otherNum;

    @ExcelProperty({"其他", "标的"})
    private String otherTarget;








    public static void main(String[] args) {
        // 文件名
        String fileName = "d://a.xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        List<CaseStatistics> list = Arrays.asList(
                new CaseStatistics("2","1","3","4","5")
        );

        EasyExcel.write(fileName, CaseTypeDTO.class)
                //sheet名称
                .sheet("模板")
                .doWrite(list);

    }

}