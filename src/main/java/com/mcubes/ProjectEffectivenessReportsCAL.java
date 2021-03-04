package com.mcubes;

import com.mcubes.model.CellData;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDoughnutChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;


class DoughnutChartData extends XDDFDoughnutChartData {

    protected DoughnutChartData(XDDFChart parent, CTDoughnutChart chart) {
        super(parent, chart);
    }
}

public class ProjectEffectivenessReportsCAL {

    private Row row;
    private Cell cell;
    private FileOutputStream outputStream;

    private List<CellData> data;

    public ProjectEffectivenessReportsCAL(List<CellData> data) {
        this.data = data;
    }

    public static void main(String[] args) {

        List<CellData> dataList = new ArrayList<>();

        CellData cellData;

        dataList.add(new CellData(3, 2, "Project Effectiveness" ,true));
        dataList.add(new CellData(4, 2, "Workspace Name:" ,true));
        dataList.add(new CellData(5, 2, "Project:" ,true));
        dataList.add(new CellData(6, 2, "Population Size:" ,true));
        dataList.add(new CellData(7, 2, "Recall Target:" ,true));
        dataList.add(new CellData(8, 2, "Probability Target Achieved @ Elusion Sample Creation:" ,true));
        dataList.add(new CellData(11, 2, "Review Results" ,true));
        dataList.add(new CellData(12, 2, "Reviewed @ Elusion Sample Creation"));
        dataList.add(new CellData(13, 2, "Found @ Elusion Sample Creation"));
        dataList.add(new CellData(14, 2, "Reviewed @ Stopping Point"));
        dataList.add(new CellData(15, 2, "Found @ Stopping Point"));
        dataList.add(new CellData(17, 2, "QC Elusion" ,true));
        dataList.add(new CellData(18, 2, "Not Reviewed @ Elusion Sample Creation"));
        dataList.add(new CellData(19, 2, "Elusion Sample Size"));
        dataList.add(new CellData(20, 2, "Relevant Docs in sample"));
        dataList.add(new CellData(21, 5, "Lower Bound" ,true, HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(21, 7, "Estimate" ,true, HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(21, 9, "Upper Bound" ,true, HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(22, 5, "%" ,true, HorizontalAlignment.CENTER));
        dataList.add(new CellData(22, 6, "      #      " ,true, HorizontalAlignment.CENTER));
        dataList.add(new CellData(22, 7, "%" ,true, HorizontalAlignment.CENTER));
        dataList.add(new CellData(22, 8, "#" ,true, HorizontalAlignment.CENTER));
        dataList.add(new CellData(22, 9, "%" ,true, HorizontalAlignment.CENTER));
        dataList.add(new CellData(22, 10, "      #      " ,true, HorizontalAlignment.CENTER));
        dataList.add(new CellData(23, 2, "How many relevant documents in UNREVIEWED @ Elusion Sample Creation?*" ,true));
        dataList.add(new CellData(24, 2, "How many relevant documents are there in total population?" ,true));
        dataList.add(new CellData(26, 5, "Lower Estimate" ,true, HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(26, 7, "Estimate" ,true, HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(26, 9, "Upper Estimate" ,true, HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(27, 2, "Estimated Recall @ Elusion Creation" ,true));
        dataList.add(new CellData(28, 2, "Estimated Recall @ Stopping Point" ,true));
        dataList.add(new CellData(29, 2, "Lower Recall Estimate = Found / (Found @ Elusion Sample Creation + Upper Bound Elusion", false, HorizontalAlignment.GENERAL, 8));
        dataList.add(new CellData(30, 2, "Recall Estimate = Found / (Found @ Elusion Sample Creation + Point Elusion", false, HorizontalAlignment.GENERAL, 8));
        dataList.add(new CellData(31, 2, "Upper Recall Estimate = Found / (Found @ Elusion Sample Creation + Lower Bound Elusion", false, HorizontalAlignment.GENERAL, 8));
        dataList.add(new CellData(33, 2, "Probability of Target Recall Calculations" ,true));
        dataList.add(new CellData(33, 8, "Estimate" ,true, HorizontalAlignment.CENTER));
        dataList.add(new CellData(34, 2, "Found @ QC Point"));
        dataList.add(new CellData(35, 2, "Recall target"));
        dataList.add(new CellData(36, 2, "Max Total True Positive (if exceeded, recall target missed)"));
        dataList.add(new CellData(37, 2, "Max False (if exceeded, recall target missed)"));
        dataList.add(new CellData(38, 2, "False Negative + True Negative (Not Reviewed @ Elusion Sample)"));
        dataList.add(new CellData(39, 2, "Max Elusion Rate (% of False Negative in Not Reviewed @ Elusion Sample"));
        dataList.add(new CellData(40, 2, "Probability Recall Target Achieved (Binomial Distribution Probability"));
        dataList.add(new CellData(41, 2, "*95% Confidence Level"));

        // TODO : replace with actual data
        dataList.add(new CellData(4, 3, "WORKSPACE NAME", true, null, 4));
        dataList.add(new CellData(5, 3, "PROJECT NAME", true, null, 4));
        dataList.add(new CellData(6, 3, 849000, true));

        cellData = new CellData(7, 3, 80, true);
        cellData.setFormat("#\"%\"");
        dataList.add(cellData);

        cellData = new CellData(8, 3, 100, true);
        cellData.setFormat("#\"%\"");
        dataList.add(cellData);

        dataList.add(new CellData(12, 3, 300000));
        dataList.add(new CellData(13, 3, 212000));
        dataList.add(new CellData(14, 3, 310000));
        dataList.add(new CellData(15, 3, 215000));

        dataList.add(new CellData(18, 3, "C6-C12", null));

        dataList.add(new CellData(19, 3, 385));
        dataList.add(new CellData(20, 3, 20));

        dataList.add(new CellData(23, 5, 3.20, null, "##.##\"%\""));
        dataList.add(new CellData(23, 6, "(E23*C18)/100", null));
        dataList.add(new CellData(23, 7, "(C20/C19)*100", "##.##\"%\""));
        dataList.add(new CellData(23, 8, "(G23*C18)/100", null));
        dataList.add(new CellData(23, 9, 7.91, null, "##.##\"%\""));
        dataList.add(new CellData(23, 10, "(I23*C18)/100", null));

        dataList.add(new CellData(24, 5, "F23+C13", null, HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(24, 7, "H23+C13", null, HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(24, 9, "J23+C13", null, HorizontalAlignment.CENTER, 1));

        dataList.add(new CellData(27, 5, "(C13/I24)*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(27, 7, "(C13/G24)*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(27, 9, "(C15/E24)*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1));

        dataList.add(new CellData(28, 5, "(C15/(C13+J23))*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(28, 7, "(C15/(C13+H23))*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1));
        dataList.add(new CellData(28, 9, "(C15/(C13+F23))*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1));

        dataList.add(new CellData(34, 8, "C13", null));
        dataList.add(new CellData(35, 8, "C7", "##\"%\""));
        dataList.add(new CellData(36, 8, "ROUNDUP(((1+H34)/H35)*100, 0)", null));
        dataList.add(new CellData(37, 8, "H36-H34", null));
        dataList.add(new CellData(38, 8, "C18", null));
        dataList.add(new CellData(39, 8, "(H37/H38)*100", "##.##\"%\""));

        ProjectEffectivenessReportsCAL obj = new ProjectEffectivenessReportsCAL(dataList);
        obj.createReport();
    }

    public void createReport() {

        try {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("sheet_01");

            for (CellData cellInfo : data) {
                setDataToCell(sheet, cellInfo);
            }

            XDDFDataSource<String> dataSourceName = XDDFDataSourcesFactory.fromArray(new String[]{"Recall", "Elusion"});
            XDDFNumericalDataSource<Integer> dataSource = XDDFDataSourcesFactory.fromArray(new Integer[]{88,12});

            drawDonutChart(sheet, 4, 7, 10, 18, dataSourceName, dataSource);

            setCellColor(sheet);

            outputStream = new FileOutputStream("files/legility.xlsx");
            workbook.write(outputStream);
            workbook.close();

        }catch (Exception e){
            e.printStackTrace();
        }

    }

    public void setDataToCell(Sheet sheet, CellData data) {

        Workbook workbook = sheet.getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        Row row = sheet.getRow(data.getRow());
        row = row == null ? sheet.createRow(data.getRow()) : row;
        Cell cell = row.getCell(data.getColumn());
        cell = cell == null ? row.createCell(data.getColumn()) : cell;
        Font font = workbook.createFont();
        font.setBold(data.isBold());
        cellStyle.setFont(font);
        if (data.getVerticalAlignment()!=null)
            cellStyle.setVerticalAlignment(data.getVerticalAlignment());
        if (data.getHorizontalAlignment()!=null)
            cellStyle.setAlignment(data.getHorizontalAlignment());
        if (data.getFormat() != null) {
            DataFormat format = workbook.createDataFormat();
            cellStyle.setDataFormat(format.getFormat(data.getFormat()));
        }
        cell.setCellStyle(cellStyle);
        switch (data.getDataType()) {
            case DOUBLE:
                cell.setCellValue(data.getDoubleValue());
                break;
            case STRING:
                cell.setCellValue(data.getStringValue());
                break;
            case BOOLEAN:
                cell.setCellValue(data.getBooleanValue());
                break;
            default:
                cell.setBlank();
        }
        sheet.autoSizeColumn(data.getColumn());
        if (data.getRowspan() !=0 || data.getColspan() != 0) {
            sheet.addMergedRegion(new CellRangeAddress(data.getRow(), data.getRow() + data.getRowspan() ,
                    data.getColumn(), data.getColumn() + data.getColspan()));
        }
        if (data.getFormula() != null)
            cell.setCellFormula(data.getFormula());
    }




    // Cell background color testing
    private void setCellColor(XSSFSheet sheet) {

        XSSFWorkbook workbook = sheet.getWorkbook();

        CellStyle style = workbook.createCellStyle();

        HSSFWorkbook hwb = new HSSFWorkbook();
        HSSFPalette palette = hwb.getCustomPalette();
        HSSFColor color = palette.findSimilarColor(132, 0, 250);
        style.setFillForegroundColor(color.getIndex());
       // HSSFColor color1 = new HSSFColor(10,1, new java.awt.Color(132, 0, 250));

        //style.setFillForegroundColor(color1.getIndex());

        System.out.println(color.getIndex());

        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFFont font = workbook.createFont();
        font.setColor(IndexedColors.WHITE.index);
        style.setFont(font);

        sheet.getRow(39).getCell(1).setCellStyle(style);
        sheet.getRow(7).getCell(1).setCellStyle(style);

        style = workbook.createCellStyle();
        hwb = new HSSFWorkbook();
        palette = hwb.getCustomPalette();
        color = palette.findSimilarColor(159, 159, 155);
        style.setFillForegroundColor(color.getIndex());

        for (int i=22; i<24; i++) {
            for (int j=1; j<10; j++){
                row = sheet.getRow(i);
                row = row==null ? sheet.createRow(i) : row;
                cell  = row.getCell(j);
                cell = cell==null ? row.createCell(j) : cell;
                cell.setCellStyle(style);
            }
        }

    }

    private void drawDonutChart(XSSFSheet sheet, int startColumn, int startRow, int endColumn, int endRow,
                                XDDFDataSource<String> dataSourceName, XDDFNumericalDataSource<Integer> dataSource){

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, startColumn, startRow, endColumn, endRow);
        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("Recall Estimate");
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);

        CTDoughnutChart doughnutChart = chart.getCTChart().addNewPlotArea().addNewDoughnutChart();
        doughnutChart.addNewVaryColors().setVal(true);
        doughnutChart.addNewHoleSize().setVal((short)50);

        XDDFDoughnutChartData data = new DoughnutChartData(chart, doughnutChart);
        //data.setVaryColors(true);
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
