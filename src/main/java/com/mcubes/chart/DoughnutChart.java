package com.mcubes.chart;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDoughnutChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;

public class DoughnutChart extends Chart{

    private int dataSourceNameFirstRow;
    private int dataSourceNameLastRow;
    private int dataSourceNameFirstColumn;
    private int dataSourceNameLastColumn;

    private int dataSourceFirstRow;
    private int dataSourceLastRow;
    private int dataSourceFirstColumn;
    private int dataSourceLastColumn;

    public DoughnutChart(String title, int startColumn, int startRow, int colspan, int rowspan, int dataSourceNameFirstRow,
                         int dataSourceNameLastRow, int dataSourceNameFirstColumn, int dataSourceNameLastColumn,
                         int dataSourceFirstRow, int dataSourceLastRow, int dataSourceFirstColumn, int dataSourceLastColumn) {
        super(title, startColumn, startRow, colspan, rowspan);
        this.dataSourceNameFirstRow = dataSourceNameFirstRow;
        this.dataSourceNameLastRow = dataSourceNameLastRow;
        this.dataSourceNameFirstColumn = dataSourceNameFirstColumn;
        this.dataSourceNameLastColumn = dataSourceNameLastColumn;
        this.dataSourceFirstRow = dataSourceFirstRow;
        this.dataSourceLastRow = dataSourceLastRow;
        this.dataSourceFirstColumn = dataSourceFirstColumn;
        this.dataSourceLastColumn = dataSourceLastColumn;
    }

    private static class DoughnutChartData extends XDDFDoughnutChartData {
        protected DoughnutChartData(XDDFChart parent, CTDoughnutChart chart) {
            super(parent, chart);
        }
    }

    @Override
    public void draw(XSSFSheet sheet) {


        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, super.getStartColumn(),
                super.getStartRow(), super.getStartColumn() + super.getColspan(), super.getStartRow() + super.getRowspan());
        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(super.getTitle());
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);

        CTDoughnutChart doughnutChart = chart.getCTChart().addNewPlotArea().addNewDoughnutChart();
        doughnutChart.addNewVaryColors().setVal(true);
        doughnutChart.addNewHoleSize().setVal((short)50);

        XDDFDoughnutChartData data = new DoughnutChartData(chart, doughnutChart);
        //data.setVaryColors(true);

        XDDFDataSource<String> dataSourceName = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(
                dataSourceNameFirstRow, dataSourceNameLastRow, dataSourceNameFirstColumn, dataSourceNameLastColumn));

        XDDFNumericalDataSource<Double> dataSource = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(
                dataSourceFirstRow, dataSourceLastRow, dataSourceFirstColumn, dataSourceLastColumn));

        XDDFChartData.Series series = data.addSeries(dataSourceName, dataSource);

        byte[][] colors = new byte[][] {
                new byte[] {(byte)132,10, (byte)250},
                new byte[] {(byte)159, (byte)159, (byte)155}
        };
        for (int i = 0; i < colors.length; i++) {
            CTPieSer ser = doughnutChart.getSerArray(0);
            ser.addNewDPt().addNewIdx().setVal(i);
            ser.getDPtArray(i).addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(colors[i]);
            CTDLbls lbl = ser.addNewDLbls();
            lbl.addNewShowVal().setVal(true);
            lbl.addNewNumFmt().setFormatCode("##\"%\"");
        }
        chart.plot(data);


    }
}
