package com.mcubes.model;

public class PieChart extends BasicChart {

    public PieChart(String title, int startCol, int startRow, int endCol, int endRow, boolean _3D) {
        super(title, startCol, startRow, endCol, endRow, _3D);
    }

    @Override
    public String toString() {
        return "PieChart" + super.toString();
    }
}
