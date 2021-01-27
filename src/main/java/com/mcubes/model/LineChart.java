package com.mcubes.model;

public class LineChart extends BasicChart {

    public LineChart(String title, int startCol, int startRow, int endCol, int endRow, boolean _3D) {
        super(title, startCol, startRow, endCol, endRow, _3D);
    }

    @Override
    public String toString() {
        return "LineChart" + super.toString();
    }
}
