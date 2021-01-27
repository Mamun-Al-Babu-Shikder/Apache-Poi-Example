package com.mcubes.writer;

import com.mcubes.model.PieChart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.List;

public class ExcelWriter {

    private enum DataType {
        BOOLEAN, STRING, INTEGER, DOUBLE, DATE
    }

    private FileOutputStream fos;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private Row row;
    private Cell cell;
    private String[] headers;
    private List<Boolean[]> booleanData;
    private List<String[]> stringData;
    private List<Integer[]> integerData;
    private List<Double[]> doubleData;
    private List<Date[]> dateData;
    private DataType selectedType;

    private PieChart pieChart;


    public ExcelWriter(File file){
        this(file, null);
    }

    public ExcelWriter(String path, String name){
        this(new File(path + File.separator + name), null);
    }

    public ExcelWriter(String path, String name, String sheetName){
        this(new File(path + File.separator + name), sheetName);
    }

    public ExcelWriter(File file, String sheetName){
        try{
            fos = new FileOutputStream(file);
            this.workbook = new XSSFWorkbook();
            if (sheetName!=null && sheetName.trim().length()!=0)
                this.sheet = workbook.createSheet(sheetName);
            else
                this.sheet = workbook.createSheet();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public void setHeadersAsArray(String[] headers){
        this.headers = headers;
    }

    public void setHeaders(String... headers){
        this.headers = headers;
    }

    public void setHeadersAsList(List<String> headers){
        if(headers!=null) {
            this.headers = (String[]) headers.toArray();
        }
    }


    public void setBooleanData(List<Boolean[]> booleanData) {
        this.booleanData = booleanData;
        this.selectedType = DataType.BOOLEAN;
    }

    public void setStringData(List<String[]> stringData) {
        this.stringData = stringData;
        this.selectedType = DataType.STRING;
    }

    public void setIntegerData(List<Integer[]> integerData) {
        this.integerData = integerData;
        this.selectedType = DataType.INTEGER;
    }

    public void setDoubleData(List<Double[]> doubleData) {
        this.doubleData = doubleData;
        this.selectedType = DataType.DOUBLE;
    }

    public void setDateData(List<Date[]> dateData) {
        this.dateData = dateData;
        this.selectedType = DataType.DATE;
    }

    public void write() throws IOException {
        this.createHeadersRow();
        this.setData();
        if (this.pieChart != null)
            createPieChart();
        this.workbook.write(fos);
        this.workbook.close();
    }

    private void createHeadersRow(){
        if(headers!=null) {
            try {
                row = sheet.createRow(0);
                for (int i = 0; i < headers.length; i++) {
                    cell = row.createCell(i);
                    cell.setCellValue(headers[i]);
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private void setData(){
        if (selectedType!=null) {
            try {
                switch (selectedType) {
                    case BOOLEAN:
                        for (int i = 0; i < booleanData.size(); i++) {
                            row = sheet.createRow(i + 1);
                            Boolean[] dataRow = booleanData.get(i);
                            for (int j = 0; j < dataRow.length; j++) {
                                cell = row.createCell(j);
                                cell.setCellValue(dataRow[j]);
                            }
                        }
                        break;
                    case STRING:
                        for (int i = 0; i < stringData.size(); i++) {
                            row = sheet.createRow(i + 1);
                            String[] dataRow = stringData.get(i);
                            for (int j = 0; j < dataRow.length; j++) {
                                cell = row.createCell(j);
                                cell.setCellValue(dataRow[j]);
                            }
                        }
                        break;
                    case INTEGER:
                        for (int i = 0; i < integerData.size(); i++) {
                            row = sheet.createRow(i + 1);
                            Integer[] dataRow = integerData.get(i);
                            for (int j = 0; j < dataRow.length; j++) {
                                cell = row.createCell(j);
                                cell.setCellValue(dataRow[j]);
                            }
                        }
                        break;
                    case DOUBLE:
                        for (int i = 0; i < doubleData.size(); i++) {
                            row = sheet.createRow(i + 1);
                            Double[] dataRow = doubleData.get(i);
                            for (int j = 0; j < dataRow.length; j++) {
                                cell = row.createCell(j);
                                cell.setCellValue(dataRow[j]);
                            }
                        }
                        break;
                    case DATE:
                        for (int i = 0; i < dateData.size(); i++) {
                            row = sheet.createRow(i + 1);
                            Date[] dataRow = dateData.get(i);
                            for (int j = 0; j < dataRow.length; j++) {
                                cell = row.createCell(j);
                                cell.setCellValue(dataRow[j]);
                            }
                        }
                        break;
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public void createPieChart(String title, boolean _3D){
        this.pieChart = new PieChart(title, 0, 0, 10, 20, _3D);
    }

    public void createPieChart(String title, int startCol, int startRow, int endCol, int endRow, boolean _3D){
        this.pieChart = new PieChart(title, startCol, startRow, endCol, endRow, _3D);
    }

    private void createPieChart(){
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, pieChart.getStartCol(), pieChart.getStartRow(), pieChart.getEndCol(), pieChart.getEndRow());
        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(pieChart.getTitle());
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);
        XDDFDataSource<String> categoryDataSource = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, headers.length-1));
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, headers.length-1));
        XDDFChartData data = chart.createData(pieChart.is_3D() ? ChartTypes.PIE3D : ChartTypes.PIE, null, null);
        data.setVaryColors(true);
        data.addSeries(categoryDataSource, values);
        chart.plot(data);
    }

}
