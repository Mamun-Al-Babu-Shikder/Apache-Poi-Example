package com.mcubes.model;

import java.util.ArrayList;
import java.util.List;

public class PieChart extends BasicChart {

    public PieChart(String title, int startCol, int startRow, int endCol, int endRow, boolean _3D) {
        super(title, startCol, startRow, endCol, endRow, _3D);

        System.out.println("notify".getClass().getSimpleName().equals("String"));
    }

    @Override
    public String toString() {
        return "PieChart" + super.toString();
    }
}
