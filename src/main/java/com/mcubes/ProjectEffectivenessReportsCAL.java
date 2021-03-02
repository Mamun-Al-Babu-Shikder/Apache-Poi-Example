package com.mcubes;

import com.mcubes.model.ProjectEffectivenessReportData;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
//import org.apache.poi.sl.draw.binding.CTSRgbColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.xmlbeans.SchemaType;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDoughnutChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.impl.CTCatAxImpl;
import org.openxmlformats.schemas.drawingml.x2006.chart.impl.CTPlotAreaImpl;
import org.openxmlformats.schemas.drawingml.x2006.main.CTScRgbColor;

import java.awt.color.ColorSpace;
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

    private List<String> dataList;
    private ProjectEffectivenessReportData data;

    public ProjectEffectivenessReportsCAL(ProjectEffectivenessReportData data) {

        this.data = data;

        dataList = new ArrayList<>();
        dataList.add("3,2,Project Effectiveness,b");
        dataList.add("4,2,Workspace Name:,b");
        dataList.add("5,2,Project:,b");
        dataList.add("6,2,Population Size:,b");
        dataList.add("7,2,Recall Target:,b");
        dataList.add("8,2,Probability Target Achieved @ Elusion Sample Creation:,b");
        dataList.add("11,2,Review Results,b");
        dataList.add("12,2,Reviewed @ Elusion Sample Creation");
        dataList.add("13,2,Found @ Elusion Sample Creation");
        dataList.add("14,2,Reviewed @ Stopping Point");
        dataList.add("15,2,Found @ Stopping Point");
        dataList.add("17,2,QC Elusion,b");
        dataList.add("18,2,Not Reviewed @ Elusion Sample Creation");
        dataList.add("19,2,Elusion Sample Size");
        dataList.add("20,2,Relevant Docs in sample");

        dataList.add("21,5,Lower Bound,b,c");
        dataList.add("21,7,Estimate,b,c");
        dataList.add("21,9,Upper Bound,b,c");

        dataList.add("22,5,%,b,c");
        dataList.add("22,6,#,b,c");
        dataList.add("22,7,%,b,c");
        dataList.add("22,8,#,b,c");
        dataList.add("22,9,%,b,c");
        dataList.add("22,10,#,b,c");

        dataList.add("23,2,How many relevant documents in UNREVIEWED @ Elusion Sample Creation?*,b");
        dataList.add("24,2,How many relevant documents are there in total population?,b");

        dataList.add("26,5,Lower Bound,b,c");
        dataList.add("26,7,Estimate,b,c");
        dataList.add("26,9,Upper Bound,b,c");

        dataList.add("27,2,Estimated Recall @ Elusion Creation,b");
        dataList.add("28,2,Estimated Recall @ Stopping Point,b");

        dataList.add("29,2,Lower Recall Estimate = Found / (Found @ Elusion Sample Creation + Upper Bound Elusion");
        dataList.add("30,2,Recall Estimate = Found / (Found @ Elusion Sample Creation + Point Elusion");
        dataList.add("31,2,Upper Recall Estimate = Found / (Found @ Elusion Sample Creation + Lower Bound Elusion");

        dataList.add("33,2,Probability of Target Recall Calculations,b");
        dataList.add("33,8,Estimate,b,c");
        dataList.add("34,2,Found @ QC Point");
        dataList.add("35,2,Recall target");
        dataList.add("36,2,Max Total True Positive (if exceeded. recall target missed)");
        dataList.add("37,2,Max False (if exceeded. recall target missed)");
        dataList.add("38,2,False Negative + True Negative (Not Reviewed @ Elusion Sample)");
        dataList.add("39,2,Max Elusion Rate (% of False Negative in Not Reviewed @ Elusion Sample");
        dataList.add("40,2,Probability Recall Target Achieved (Binomial Distribution Probability");
        dataList.add("41,2,*95% Confidence Level");

    }

    public static void main(String[] args) {

        ProjectEffectivenessReportData data = new ProjectEffectivenessReportData();
        data.setWorkspaceName("WORKSPACE NAME");
        data.setProjectName("PROJECT NAME");
        data.setPopulationSize(849000);
        data.setRecallTarget(80);
        data.setProbabilityTargetAchieved(100);
        data.setReviewedElusionSampleCreation(300000);
        data.setFoundElusionSampleCreation(212000);
        data.setReviewedStoppingPoint(310000);
        data.setFoundStoppingPoint(215000);
        data.setElusionSampleSize(385);
        data.setRelevantDocsInSample(20);

        ProjectEffectivenessReportsCAL obj = new ProjectEffectivenessReportsCAL(data);
        obj.createReport();

    }

    public void createReport() {

        try {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("sheet_01");


            row = sheet.createRow(3);
            cell = row.createCell(1);



           // dataList.forEach(v->setCellValue(sheet, v));

            for (String data : dataList) {
                boolean isBold = false;
                HorizontalAlignment alignment = HorizontalAlignment.GENERAL;
                String values[] = data.split(",");
                if (values.length > 3 && values[3].equals("b")) isBold = true;
                if (values.length > 4) {
                    switch (values[4]){
                        case "c":
                            alignment = HorizontalAlignment.CENTER;
                            break;
                        case "r":
                            alignment = HorizontalAlignment.RIGHT;
                            break;
                    }
                }
                setStyleAbleValue(sheet, Integer.parseInt(values[0]), Integer.parseInt(values[1]), values[2],
                        isBold, alignment, null);
                sheet.autoSizeColumn(Integer.parseInt(values[1])-1);
            }


            mergeCell(sheet);

            CellStyle style = workbook.createCellStyle();
            style.setFillForegroundColor(IndexedColors.VIOLET.index);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            XSSFFont font = workbook.createFont();
            font.setColor(IndexedColors.WHITE.index);
            style.setFont(font);
            sheet.getRow(39).getCell(1).setCellStyle(style);
            sheet.getRow(7).getCell(1).setCellStyle(style);



            for (short i=0; i<120; i++){
                CellStyle style2 = workbook.createCellStyle();
                style2.setFillForegroundColor(i);
                // style.setFillForegroundColor(IndexedColors.VIOLET.index);
                style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                row = sheet.createRow(i);
                cell = row.createCell(0);
                cell.setCellValue("hello");
                cell.setCellStyle(style2);
            }

            //style.setFillBackgroundColor((short)new java.awt.Color(0,0,0).getRGB());
           // sheet.getRow(39).getCell(1).setCellStyle(style);

          //  XSSFWorkbook workbook2 = new XSSFWorkbook();
           // XSSFColor color = new XSSFColor(workbook2.getStylesSource().getIndexedColors());
            //color.setARGBHex("#000000".substring(1));

            //    XSSFCellStyle style = new XSSFCellStyle();

            //cell.setCellStyle();

            setFormula(sheet);

            drawDonutChart(sheet);

            setData(sheet);

            outputStream = new FileOutputStream("files/legility.xlsx");
            workbook.write(outputStream);
            workbook.close();

        }catch (Exception e){
            e.printStackTrace();
        }

    }


    /*
    private void setCellValue(Sheet sheet, String data) {
        String[] values = data.split(",");
        Row row = sheet.getRow(Integer.parseInt(values[0])-1);
        row = row == null ? sheet.createRow(Integer.parseInt(values[0])-1) : row;
        Cell cell = row.createCell(Integer.parseInt(values[1])-1);
        cell.setCellValue(values[2]);
        sheet.autoSizeColumn(Integer.parseInt(values[1])-1);
    }

    private CellStyle getStyleForFont(Workbook workbook){
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        Font font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        return cellStyle;
    }
     */



    private void setStyleAbleValue(Sheet sheet, int rowNumber, int colNumber, String value, boolean boldText,
                                   HorizontalAlignment alignment, String fontName){
        Workbook workbook = sheet.getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        Row row = sheet.getRow(rowNumber-1);
        row = row == null ? sheet.createRow(rowNumber-1) : row;
        Cell cell = row.createCell(colNumber-1);
        Font font = workbook.createFont();
        font.setBold(boldText);
        if (fontName!=null)
            font.setFontName(fontName);
        if (alignment!=null)
            cellStyle.setAlignment(alignment);
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(value);
    }

    private void mergeCell(XSSFSheet sheet) {
        sheet.addMergedRegion(CellRangeAddress.valueOf("E21:F21"));
        sheet.addMergedRegion(CellRangeAddress.valueOf("G21:H21"));
        sheet.addMergedRegion(CellRangeAddress.valueOf("I21:J21"));

        for (int i=24; i<29; i++) {
            sheet.addMergedRegion(CellRangeAddress.valueOf("E"+i+":F"+i));
            sheet.addMergedRegion(CellRangeAddress.valueOf("G"+i+":H"+i));
            sheet.addMergedRegion(CellRangeAddress.valueOf("I"+i+":J"+i));
        }

        sheet.addMergedRegion(CellRangeAddress.valueOf("B29:Z29"));
        sheet.addMergedRegion(CellRangeAddress.valueOf("B30:Z30"));
        sheet.addMergedRegion(CellRangeAddress.valueOf("B31:Z31"));

        for (int i=34; i<42; i++)
            sheet.addMergedRegion(CellRangeAddress.valueOf("B"+i+":E"+i));

    }

    private void setFormula(Sheet sheet) {

        sheet.getRow(22).createCell(6).setCellFormula("(C20/C19)*100");
        //cell.setCellType(CellType.FORMULA);
        sheet.getRow(22).createCell(7).setCellFormula("G23*C18");
        sheet.getRow(22).createCell(5).setCellFormula("(E23*C18)/100");
        sheet.getRow(22).createCell(9).setCellFormula("(I23*C18)/100");

        sheet.getRow(23).createCell(4).setCellFormula("F23+C18");
        sheet.getRow(23).createCell(6).setCellFormula("H23+C18");
        sheet.getRow(23).createCell(8).setCellFormula("J23+C18");
        sheet.getRow(17).createCell(2).setCellFormula("C6-C12");

    }

    private void drawDonutChart(XSSFSheet sheet){

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 7, 10, 18);
        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("Recall Estimate");
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);

        XDDFDataSource<String> dataSourceName = XDDFDataSourcesFactory.fromArray(new String[]{"Recall", "Elusion"});
        XDDFNumericalDataSource<Integer> dataSource = XDDFDataSourcesFactory.fromArray(new Integer[]{88,12});

        CTDoughnutChart doughnutChart = chart.getCTChart().addNewPlotArea().addNewDoughnutChart();
        doughnutChart.addNewVaryColors().setVal(true);
        doughnutChart.addNewHoleSize().setVal((short)50);

        XDDFDoughnutChartData data = new DoughnutChartData(chart, doughnutChart);
        //data.setVaryColors(true);
        XDDFChartData.Series series = data.addSeries(dataSourceName, dataSource);

        byte[][] colors = new byte[][] {
                new byte[] {(byte)132,0, (byte)250},
                new byte[] {(byte)159, (byte)159, (byte)155}
        };

        for (int i = 0; i < 2; i++) {
            CTPieSer ser = doughnutChart.getSerArray(0);
            ser.addNewDPt().addNewIdx().setVal(i);
            ser.getDPtArray(i).addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(colors[i]);
        }
        chart.plot(data);
    }

    private void setData(XSSFSheet sheet) {

        sheet.getRow(3).createCell(2).setCellValue(data.getWorkspaceName());
        sheet.getRow(4).createCell(2).setCellValue(data.getProjectName());
        sheet.getRow(5).createCell(2).setCellValue(data.getPopulationSize());
        sheet.getRow(6).createCell(2).setCellValue(data.getRecallTarget());
        sheet.getRow(7).createCell(2).setCellValue(data.getProbabilityTargetAchieved());
        sheet.getRow(11).createCell(2).setCellValue(data.getReviewedElusionSampleCreation());
        sheet.getRow(12).createCell(2).setCellValue(data.getFoundElusionSampleCreation());
        sheet.getRow(13).createCell(2).setCellValue(data.getReviewedStoppingPoint());
        sheet.getRow(14).createCell(2).setCellValue(data.getFoundStoppingPoint());
        sheet.getRow(18).createCell(2).setCellValue(data.getElusionSampleSize());
        sheet.getRow(19).createCell(2).setCellValue(data.getRelevantDocsInSample());

    }


}
