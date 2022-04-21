package com.cw.backlinxian.vo;

public class DailyCountVo {
    private String date;
    private int type1Count;
    private int type2Count;
    private int type3Count;
    private int totalCount;
    private int homeCount;
    private int groupCount;
    private int jianceCount;

    @Override
    public String toString() {
        return "DailyCountVo{" +
                "date='" + date + '\'' +
                ", type1Count=" + type1Count +
                ", type2Count=" + type2Count +
                ", type3Count=" + type3Count +
                ", totalCount=" + totalCount +
                ", homeCount=" + homeCount +
                ", groupCount=" + groupCount +
                ", jianceCount=" + jianceCount +
                '}';
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public int getType1Count() {
        return type1Count;
    }

    public void setType1Count(int type1Count) {
        this.type1Count = type1Count;
    }

    public int getType2Count() {
        return type2Count;
    }

    public void setType2Count(int type2Count) {
        this.type2Count = type2Count;
    }

    public int getType3Count() {
        return type3Count;
    }

    public void setType3Count(int type3Count) {
        this.type3Count = type3Count;
    }

    public int getTotalCount() {
        return totalCount;
    }

    public void setTotalCount(int totalCount) {
        this.totalCount = totalCount;
    }

    public int getHomeCount() {
        return homeCount;
    }

    public void setHomeCount(int homeCount) {
        this.homeCount = homeCount;
    }

    public int getGroupCount() {
        return groupCount;
    }

    public void setGroupCount(int groupCount) {
        this.groupCount = groupCount;
    }

    public int getJianceCount() {
        return jianceCount;
    }

    public void setJianceCount(int jianceCount) {
        this.jianceCount = jianceCount;
    }
}
