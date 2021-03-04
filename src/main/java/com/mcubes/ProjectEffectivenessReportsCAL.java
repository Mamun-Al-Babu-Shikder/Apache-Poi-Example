package com.mcubes;

import com.mcubes.chart.DoughnutChart;
import com.mcubes.model.CellData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDoughnutChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;

import java.awt.Color;
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

        dataList.add(new CellData(3, 2, "Project Effectiveness" ,true, null, 0, null, null, null));
        dataList.add(new CellData(4, 2, "Workspace Name:" ,true, null, 0, null, null, null));
        dataList.add(new CellData(5, 2, "Project:" ,true, null, 0, null, null, null));
        dataList.add(new CellData(6, 2, "Population Size:" ,true, null, 0, null, null, null));
        dataList.add(new CellData(7, 2, "Recall Target:" ,true, null, 0, null, null, null));
        dataList.add(new CellData(8, 2, "Probability Target Achieved @ Elusion Sample Creation:" , true, HorizontalAlignment.GENERAL, 0, Color.WHITE, new Color(0x8400FA), null));
        dataList.add(new CellData(11, 2, "Review Results" ,true, null, 0, null, null, null));
        dataList.add(new CellData(12, 2, "Reviewed @ Elusion Sample Creation", false, null, 0, null, null, null));
        dataList.add(new CellData(13, 2, "Found @ Elusion Sample Creation", false , null, 0, null, null, null));
        dataList.add(new CellData(14, 2, "Reviewed @ Stopping Point", false, null, 0, null, null, null));
        dataList.add(new CellData(15, 2, "Found @ Stopping Point", false, null, 0, null, null, null));
        dataList.add(new CellData(17, 2, "QC Elusion" ,true, null, 0, null, null, null));
        dataList.add(new CellData(18, 2, "Not Reviewed @ Elusion Sample Creation", false, null, 0, null, null, null));
        dataList.add(new CellData(19, 2, "Elusion Sample Size", false, null, 0, null, null, null));
        dataList.add(new CellData(20, 2, "Relevant Docs in sample", false, null, 0, null, null, null));
        dataList.add(new CellData(21, 5, "Lower Bound" ,true, HorizontalAlignment.CENTER, 1, null, null, null));
        dataList.add(new CellData(21, 7, "Estimate" ,true, HorizontalAlignment.CENTER, 1, Color.BLACK, new Color(0xD276FF), null));
        dataList.add(new CellData(21, 9, "Upper Bound" ,true, HorizontalAlignment.CENTER, 1, null, null, null));
        dataList.add(new CellData(22, 5, "%" ,true, HorizontalAlignment.CENTER, 0, null, null, null));
        dataList.add(new CellData(22, 6, "      #      " ,true, HorizontalAlignment.CENTER, 0, null, null, null));
        dataList.add(new CellData(22, 7, "%" ,true, HorizontalAlignment.CENTER, 0, Color.BLACK, new Color(0xD276FF), null));
        dataList.add(new CellData(22, 8, "#" ,true, HorizontalAlignment.CENTER, 0, Color.BLACK, new Color(0xD276FF), null));
        dataList.add(new CellData(22, 9, "%" ,true, HorizontalAlignment.CENTER, 0, null, null, null));
        dataList.add(new CellData(22, 10, "      #      " ,true, HorizontalAlignment.CENTER, 0, null, null, null));
        dataList.add(new CellData(23, 2, "How many relevant documents in UNREVIEWED @ Elusion Sample Creation?*" ,true, null, 0, null, null, null));
        dataList.add(new CellData(24, 2, "How many relevant documents are there in total population?" ,true, null, 0, null, null, null));
        dataList.add(new CellData(26, 5, "Lower Estimate" ,true, HorizontalAlignment.CENTER, 1, null, null, null));
        dataList.add(new CellData(26, 7, "Estimate" ,true, HorizontalAlignment.CENTER, 1, Color.BLACK, new Color(0xD276FF), null));
        dataList.add(new CellData(26, 9, "Upper Estimate" ,true, HorizontalAlignment.CENTER, 1, null, null, null));
        dataList.add(new CellData(27, 2, "Estimated Recall @ Elusion Creation" ,true, null, 0, null, null, null));
        dataList.add(new CellData(28, 2, "Estimated Recall @ Stopping Point" ,true, null, 0, null, null, null));
        dataList.add(new CellData(29, 2, "Lower Recall Estimate = Found / (Found @ Elusion Sample Creation + Upper Bound Elusion", false, HorizontalAlignment.GENERAL, 8, null, null, null));
        dataList.add(new CellData(30, 2, "Recall Estimate = Found / (Found @ Elusion Sample Creation + Point Elusion", false, HorizontalAlignment.GENERAL, 8, null, null, null));
        dataList.add(new CellData(31, 2, "Upper Recall Estimate = Found / (Found @ Elusion Sample Creation + Lower Bound Elusion", false, HorizontalAlignment.GENERAL, 8, null, null, null));
        dataList.add(new CellData(33, 2, "Probability of Target Recall Calculations" ,true, null, 0, null, null, null));
        dataList.add(new CellData(33, 8, "Estimate" ,true, HorizontalAlignment.CENTER, 0, null, null, null));
        dataList.add(new CellData(34, 2, "Found @ QC Point", false , null, 0, null, null, null));
        dataList.add(new CellData(35, 2, "Recall target", false, null, 0, null, null, null));
        dataList.add(new CellData(36, 2, "Max Total True Positive (if exceeded, recall target missed)", false, null, 0, null, null, null));
        dataList.add(new CellData(37, 2, "Max False (if exceeded, recall target missed)", false, null, 0, null, null, null));
        dataList.add(new CellData(38, 2, "False Negative + True Negative (Not Reviewed @ Elusion Sample)", false, null, 0, null, null, null));
        dataList.add(new CellData(39, 2, "Max Elusion Rate (% of False Negative in Not Reviewed @ Elusion Sample", false, null, 0, null, null, null));
        dataList.add(new CellData(40, 2, "Probability Recall Target Achieved (Binomial Distribution Probability", true, HorizontalAlignment.GENERAL, 2, Color.WHITE, new Color(0x8400FA), null));
        dataList.add(new CellData(41, 2, "*95% Confidence Level", false, null, 0, null, null, null));

        // TODO : replace with actual data

        dataList.add(new CellData(4, 3, "WORKSPACE NAME", true, null, 4, null, null, null));
        dataList.add(new CellData(5, 3, "PROJECT NAME", true, null, 4, null, null, null));
        dataList.add(new CellData(6, 3, Long.toString(849000), true, null, 0, null, null, null));
        dataList.add(new CellData(7, 3, Integer.toString(80), true, null, 0, null, null, "#\"%\""));
        dataList.add(new CellData(8, 3, Integer.toString(100), true, null, 0, null, null, "#\"%\""));

        dataList.add(new CellData(12, 3, Long.toString(300000), false, null, 0, null, null, null));
        dataList.add(new CellData(13, 3, Long.toString(212000), false, null, 0, null, null, null));
        dataList.add(new CellData(14, 3, Long.toString(310000), false, null, 0, null, null, null));
        dataList.add(new CellData(15, 3, Long.toString(215000), false, null, 0, null, null, null));

        dataList.add(new CellData(18, 3, "C6-C12", null, null, 0, false, null, null));

        dataList.add(new CellData(19, 3, Integer.toString(385), false, null, 0, null, null, null));
        dataList.add(new CellData(20, 3, Integer.toString(20), false, null, 0, null, null, null));

        dataList.add(new CellData(23, 5, Double.toString(3.20), false, null, 0, null, null, "##.##\"%\""));

        dataList.add(new CellData(23, 6, "(E23*C18)/100", null, null, 0, false, null, null));
        dataList.add(new CellData(23, 7, "(C20/C19)*100", "##.##\"%\"", null, 0, false, null, null));
        dataList.add(new CellData(23, 8, "(G23*C18)/100", null, null, 0, false, null, null));

        dataList.add(new CellData(23, 9, Double.toString(7.91), false, null, 0, null, null, "##.##\"%\""));

        dataList.add(new CellData(23, 10, "(I23*C18)/100", null, null, 0, false, null, null));

        dataList.add(new CellData(24, 5, "F23+C13", null, HorizontalAlignment.CENTER, 1, false, null, null));
        dataList.add(new CellData(24, 7, "H23+C13", null, HorizontalAlignment.CENTER, 1, false, null, null));
        dataList.add(new CellData(24, 9, "J23+C13", null, HorizontalAlignment.CENTER, 1, false, null, null));

        dataList.add(new CellData(27, 5, "(C13/I24)*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1, false, null, null));
        dataList.add(new CellData(27, 7, "(C13/G24)*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1, false, null, null));
        dataList.add(new CellData(27, 9, "(C15/E24)*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1, false, null, null));

        dataList.add(new CellData(28, 5, "(C15/(C13+J23))*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1, false, null, null));
        dataList.add(new CellData(28, 7, "(C15/(C13+H23))*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1, false, null, null));
        dataList.add(new CellData(28, 9, "(C15/(C13+F23))*100", "##.##\"%\"", HorizontalAlignment.CENTER, 1, false, null, null));

        dataList.add(new CellData(34, 8, "C13", null, null, 0, false, null, null));

        dataList.add(new CellData(35, 8, "C7", "##\"%\"", null, 0, false, null, null));
        dataList.add(new CellData(36, 8, "ROUNDUP(((1+H34)/H35)*100, 0)", null, null, 0, false, null, null));
        dataList.add(new CellData(37, 8, "H36-H34", null, null, 0, false, null, null));
        dataList.add(new CellData(38, 8, "C18", null, null, 0, false, null, null));
        dataList.add(new CellData(39, 8, "(H37/H38)*100", "##.##\"%\"", null, 0, false, null, null));


        dataList.add(new CellData(11, 8, "Recall", false, null, 0, Color.white, null, null));
        dataList.add(new CellData(12, 8, "Estimate", false, null, 0, Color.white, null, null));

        dataList.add(new CellData(11, 7, "(C13/G24)*100", null, null, 0, false, Color.white, null));
        dataList.add(new CellData(12, 7, "((G24-C13)/G24)*100", null, null, 0, false, Color.white, null));


        DoughnutChart donutChart = new DoughnutChart("Recall Estimate",4,7, 6, 11,
                10, 11, 7, 7,
                10, 11, 6, 6);
        dataList.add(new CellData(donutChart));

        ProjectEffectivenessReportsCAL obj = new ProjectEffectivenessReportsCAL(dataList);
        obj.createReport();
    }

    public void createReport() {

        try {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("sheet_01");

          //  sheet.createRow(10).createCell(7).setCellValue("Recall");
         //   sheet.createRow(11).createCell(7).setCellValue("Estimate");

           // sheet.createRow(10).createCell(6).setCellValue(88);
          //  sheet.createRow(11).createCell(6).setCellValue(12);

            for (CellData cellInfo : data) {
                setDataToCell(sheet, cellInfo);
            }



            //XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1,3,4,5));;

            //XDDFDataSource<String> dataSourceName = XDDFDataSourcesFactory.fromArray(new String[]{"Recall", "Elusion"});
            //XDDFNumericalDataSource<Integer> dataSource = XDDFDataSourcesFactory.fromArray(new Integer[]{88,12});

            //drawDonutChart(sheet, 4, 7, 10, 18, dataSourceName, dataSource);

           // setCellColor(sheet);

            outputStream = new FileOutputStream("files/legility.xlsx");
            workbook.write(outputStream);
            workbook.close();

        }catch (Exception e){
            e.printStackTrace();
        }

    }

    @SuppressWarnings("deprecation")
    public void setDataToCell(XSSFSheet sheet, CellData data) {

        if (data.getDataType() == CellData.DataType.CHART) {
            data.getChart().draw(sheet);
        } else {

            XSSFWorkbook workbook = sheet.getWorkbook();
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            Row row = sheet.getRow(data.getRow());
            row = row == null ? sheet.createRow(data.getRow()) : row;
            Cell cell = row.getCell(data.getColumn());
            cell = cell == null ? row.createCell(data.getColumn()) : cell;
            XSSFFont font = workbook.createFont();
            if (data.getFontColor() != null)
                font.setColor(new XSSFColor(data.getFontColor()));
            font.setBold(data.isBold());
            cellStyle.setFont(font);
            if (data.getCellColor() != null) {
                cellStyle.setFillBackgroundColor(new XSSFColor(data.getCellColor()));
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            if (data.getVerticalAlignment() != null)
                cellStyle.setVerticalAlignment(data.getVerticalAlignment());
            if (data.getHorizontalAlignment() != null)
                cellStyle.setAlignment(data.getHorizontalAlignment());
            if (data.getFormat() != null) {
                DataFormat format = workbook.createDataFormat();
                cellStyle.setDataFormat(format.getFormat(data.getFormat()));
            }
            cell.setCellStyle(cellStyle);
            switch (data.getDataType()) {
                case STRING:
                    cell.setCellValue(data.getStringValue());
                    break;
                default:
                    cell.setBlank();
            }
            sheet.autoSizeColumn(data.getColumn());
            if (data.getRowspan() != 0 || data.getColspan() != 0) {
                sheet.addMergedRegion(new CellRangeAddress(data.getRow(), data.getRow() + data.getRowspan(),
                        data.getColumn(), data.getColumn() + data.getColspan()));
            }
            if (data.getFormula() != null)
                cell.setCellFormula(data.getFormula());
        }
    }




    // Cell background color testing
    private void setCellColor(XSSFSheet sheet) {

        XSSFWorkbook workbook = sheet.getWorkbook();
        XSSFCellStyle style = workbook.createCellStyle();

        style.setFillBackgroundColor(new XSSFColor(new java.awt.Color(0xB622D0)));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        sheet.getRow(39).getCell(1).setCellStyle(style);


        style = workbook.createCellStyle();
        style.setFillBackgroundColor(new XSSFColor(new java.awt.Color(0xEF0B0B)));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        sheet.getRow(7).getCell(1).setCellStyle(style);












        /*
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

         */

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
