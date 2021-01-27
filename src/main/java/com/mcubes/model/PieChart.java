package com.mcubes.model;

public class PieChart {

    private String title;
    private int startCol;
    private int startRow;
    private int endCol;
    private int endRow;
    private boolean _3D;

    public PieChart(String title, int startCol, int startRow, int endCol, int endRow, boolean _3D) {
        this.title = title;
        this.startCol = startCol;
        this.startRow = startRow;
        this.endCol = endCol;
        this.endRow = endRow;
        this._3D = _3D;
    }

    public String getTitle() {
        return title;
    }

    public int getStartCol() {
        return startCol;
    }

    public int getStartRow() {
        return startRow;
    }

    public int getEndCol() {
        return endCol;
    }

    public int getEndRow() {
        return endRow;
    }

    public boolean is_3D() {
        return _3D;
    }
}
