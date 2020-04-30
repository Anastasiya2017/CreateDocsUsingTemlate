package com.test;

public class Template {
    String taskNum;
    String imgCount;
    String itemGroup;
    String itemName;
    String color;
    String interiorColor;
    String startDate;
    String endDate;
    String size;
    String mainImage;
    String resultImage;

    public Template(String taskNum, String imgCount, String itemGroup,
                    String itemName, String color, String interiorColor,
                    String startDate, String endDate, String size, String mainImage,
                    String resultImage) {
        setTaskNum(taskNum);
        setImgCount(imgCount);
        setItemGroup(itemGroup);
        setItemName(itemName);
        setColor(color);
        setInteriorColor(interiorColor);
        setStartDate(startDate);
        setEndDate(endDate);
        setSize(size);
        setMainImage(mainImage);
        setResultImage(resultImage);
    }


    public String getTaskNum() {
        return taskNum;
    }

    public void setTaskNum(String taskNum) {
        this.taskNum = taskNum;
    }

    public String getImgCount() {
        return imgCount;
    }

    public void setImgCount(String imgCount) {
        this.imgCount = imgCount;
    }

    public String getItemGroup() {
        return itemGroup;
    }

    public void setItemGroup(String itemGroup) {
        this.itemGroup = itemGroup;
    }

    public String getItemName() {
        return itemName;
    }

    public void setItemName(String itemName) {
        this.itemName = itemName;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public String getInteriorColor() {
        return interiorColor;
    }

    public void setInteriorColor(String interiorColor) {
        this.interiorColor = interiorColor;
    }

    public String getStartDate() {
        return startDate;
    }

    public void setStartDate(String startDate) {
        this.startDate = startDate;
    }

    public String getEndDate() {
        return endDate;
    }

    public void setEndDate(String endDate) {
        this.endDate = endDate;
    }

    public String getSize() {
        return size;
    }

    public void setSize(String size) {
        this.size = size;
    }

    public String getMainImage() {
        return mainImage;
    }

    public void setMainImage(String mainImage) {
        this.mainImage = mainImage;
    }

    public String getResultImage() {
        return resultImage;
    }

    public void setResultImage(String resultImage) {
        this.resultImage = resultImage;
    }
}
