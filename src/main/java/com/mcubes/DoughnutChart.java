package com.mcubes;

import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDoughnutChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;

import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;

public class DoughnutChart {

    public static void main(String[] args) throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");

        Row row;
        Cell cell;
        for (int r = 0; r < 3; r++) {
            row = sheet.createRow(r);
            cell = row.createCell(0);
            cell.setCellValue("S" + r);
            cell = row.createCell(1);
            cell.setCellValue(r+1);
        }

        XSSFDrawing drawing = (XSSFDrawing)sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 5, 20);

        XSSFChart chart = drawing.createChart(anchor);

        CTChart ctChart = ((XSSFChart)chart).getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();
        CTDoughnutChart ctDoughnutChart = ctPlotArea.addNewDoughnutChart();
        ctDoughnutChart.addNewVaryColors().setVal(true);
        ctDoughnutChart.addNewHoleSize().setVal((short)50);

        CTPieSer ctPieSer = ctDoughnutChart.addNewSer();

        //ctPieSer.addNewIdx().setVal(0);


        CTAxDataSource cttAxDataSource = ctPieSer.addNewCat();
        CTStrRef ctStrRef = cttAxDataSource.addNewStrRef();
        ctStrRef.setF("Sheet1!$A$1:$A$3");
        CTNumDataSource ctNumDataSource = ctPieSer.addNewVal();
        CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
        ctNumRef.setF("Sheet1!$B$1:$B$3");



        // Data point colors; is necessary for showing data points in Calc

        int pointCount = 3;
        for (int p = 0; p < pointCount; p++) {
            ctChart.getPlotArea().getDoughnutChartArray(0).getSerArray(0).addNewDPt().addNewIdx().setVal(p);
            ctChart.getPlotArea().getDoughnutChartArray(0).getSerArray(0).getDPtArray(p)
                    .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(p+10));
        }



        System.out.println(ctChart);

        FileOutputStream fileOut = new FileOutputStream("files/workbook.xlsx");
        wb.write(fileOut);
        fileOut.close();
        wb.close();
    }
}