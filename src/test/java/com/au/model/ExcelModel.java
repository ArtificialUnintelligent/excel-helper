package com.au.model;

import com.au.annotation.ExcelColumn;

/**
 * @author:artificialunintelligent
 * @Date:2019-07-05
 * @Time:17:55
 */
public class ExcelModel {

    @ExcelColumn(secondTitle = "aaa")
    private String remark;

    @ExcelColumn(secondTitle = "名称", firstTitle = "aaa")
    private String name;

    @ExcelColumn(secondTitle = "颜色", firstTitle = "aaa")
    private String color;

    @ExcelColumn(secondTitle = "性别", firstTitle = "bbb")
    private String sex;

    @ExcelColumn(secondTitle = "年龄", firstTitle = "bbb")
    private String age;

    @ExcelColumn(secondTitle = "人种", firstTitle = "ccc")
    private String kind;

    @ExcelColumn(secondTitle = "工作", firstTitle = "ccc")
    private String job;

    @ExcelColumn(secondTitle = "身高", firstTitle = "ccc")
    private String height;

    @ExcelColumn(secondTitle = "体重", firstTitle = "ccc")
    private String weight;



    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getKind() {
        return kind;
    }

    public void setKind(String kind) {
        this.kind = kind;
    }

    public String getJob() {
        return job;
    }

    public void setJob(String job) {
        this.job = job;
    }

    public String getHeight() {
        return height;
    }

    public void setHeight(String height) {
        this.height = height;
    }

    public String getWeight() {
        return weight;
    }

    public void setWeight(String weight) {
        this.weight = weight;
    }

    public String getRemark() {
        return remark;
    }

    public void setRemark(String remark) {
        this.remark = remark;
    }
}
