package com.mcubes;

import java.io.*;
import java.util.*;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.stream.Collectors;

import com.mcubes.writer.ExcelWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

public class Main {

    public static void main(String[] args) throws Exception {

        Integer a = 5;
        fun(a);

        System.out.println(a);
    }

    public static void fun(Integer a){
        ++a;
    }

     public static List<Double> extractedPredictionScore() throws IOException {
        List<Double> scoreList = new ArrayList<>();
        File file = new File("files/ps.csv");
        FileReader fr = new FileReader(file);
        BufferedReader br = new BufferedReader(fr);
        String line = "";
        while ((line=br.readLine())!=null){
           // System.out.println(line);
            String[] splitString = line.split(",");
            try{
                String score = splitString[2];
                if(splitString[2].matches("[-]?\\d+[\\.]\\d+")){
                    scoreList.add(Double.parseDouble(score));
                }
            }catch (IndexOutOfBoundsException e){
                System.out.println(e);
            }
        }
        br.readLine();
        fr.close();
        return scoreList;
     }


    public static void writeSimpleData(){
        ExcelWriter writer = new ExcelWriter("files", "test_excel_01.xlsx", "Sheet_no_01");
        writer.setHeaders("Bangladesh", "India", "Island", "Germany", "Brazil");
        List<Integer[]> data = new ArrayList<>();
        data.add(new Integer[]{1,2,3,4,5});
        data.add(new Integer[]{11,22,4,56,87});
        data.add(new Integer[]{15,62,37,48,5});
        data.add(new Integer[]{5,7,8,23,7});
        data.add(new Integer[]{61,82,13,54,57});
        data.add(new Integer[]{15,25,53,74,55});
        writer.setIntegerData(data);
        try {
            writer.write();
            System.out.println("Data write success!");
        }catch (IOException e){
            e.printStackTrace();
        }
    }

    public static void pieChartDemo(){
        ExcelWriter ew = new ExcelWriter(new File("files/pie_test.xlsx"));
        ew.setHeaders("Chrome", "Firefox", "Internet Explorer", "Safari", "Edge", "Opera", "Other");
        List<Double[]> data = new ArrayList<>();
        data.add(new Double[]{62.74, 10.57, 7.23, 5.58, 4.02, 1.92, 7.62});
        ew.setDoubleData(data);
        ew.createPieChart("Browser market shares. January, 2018", 0, 4, 7, 20, true);
        try {
            ew.write();
            System.out.println("Data write success!");
        }catch (IOException e){
            e.printStackTrace();
        }
    }

    public static void lineChartDemo(){
        ExcelWriter ew = new ExcelWriter("files", "line_chart_demo.xlsx", "sheet_no_01");
        ew.setHeaders("Installation", "Manufacturing", "Sales & Distribution", "Project Development", "Other");
        List<Integer[]> data = new ArrayList<>();
        data.add(new Integer[]{439, 525, 577, 695, 970});
        data.add(new Integer[]{249, 240, 294, 298, 324});
        data.add(new Integer[]{248, 146, 274, 297, 314});
        ew.setIntegerData(data);
        ew.createLineChart("Solar Employment Growth by Sector, 2010-2016",0, 10, 10, 25, false);
        try {
            ew.write();
            System.out.println("Data write success!");
        }catch (IOException e){
            e.printStackTrace();
        }
    }


    public static void barChartDemo(){








    }


    public static void barColumnChart() throws FileNotFoundException, IOException {

        try (XSSFWorkbook wb = new XSSFWorkbook()) {

            String sheetName = "CountryBarChart";//"CountryColumnChart";

            XSSFSheet sheet = wb.createSheet(sheetName);

            // Create row and put some cells in it. Rows and cells are 0 based.

            Row row = sheet.createRow((short) 0);

            Cell cell = row.createCell((short) 0);
            cell.setCellValue("Russia");

            cell = row.createCell((short) 1);
            cell.setCellValue("Canada");

            cell = row.createCell((short) 2);
            cell.setCellValue("USA");

            cell = row.createCell((short) 3);
            cell.setCellValue("China");

            cell = row.createCell((short) 4);
            cell.setCellValue("Brazil");

            cell = row.createCell((short) 5);
            cell.setCellValue("Australia");

            cell = row.createCell((short) 6);
            cell.setCellValue("India");

            row = sheet.createRow((short) 1);

            cell = row.createCell((short) 0);
            cell.setCellValue(17098242);

            cell = row.createCell((short) 1);
            cell.setCellValue(9984670);

            cell = row.createCell((short) 2);
            cell.setCellValue(9826675);

            cell = row.createCell((short) 3);
            cell.setCellValue(9596961);

            cell = row.createCell((short) 4);
            cell.setCellValue(8514877);

            cell = row.createCell((short) 5);
            cell.setCellValue(7741220);

            cell = row.createCell((short) 6);
            cell.setCellValue(3287263);







            List<Double> scoreList = extractedPredictionScore();
            long a = scoreList.stream().distinct().count();
            System.out.println(a);
            System.out.println(scoreList.size());

            Map<Double, Long> scoreMap =  scoreList.stream().collect(Collectors.groupingBy(e -> e, Collectors.counting()));

            Double[] cat = new Double[scoreMap.size()];
            Long[] val = new Long[scoreMap.size()];
            int idx = 0;
            for (Map.Entry<Double, Long> m : scoreMap.entrySet()){
                cat[idx] = m.getKey();
                val[idx] = m.getValue();
                idx++;
            }

            Arrays.sort(cat);
            Arrays.sort(val);

            System.out.println(cat[cat.length-1]+", "+val[val.length-1]);

            /*
            Row row;
            Cell cell;

            row = sheet.createRow(0);
            for (int i=1200; i<cat.length; i++){
                cell = row.createCell(i);
                cell.setCellValue(cat[i]);
            }

            row = sheet.createRow(1);
            for (int i=1200; i<val.length; i++){
                cell = row.createCell(i);
                cell.setCellValue(val[i]);
                System.out.println(val[i]);
            }

             */

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 4, 10, 35);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Area-wise Top Seven Countries");
            chart.setTitleOverlay(false);

            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.BOTTOM);

            chart.deleteLegend();

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("Country");
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("Area");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

            XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange(sheet,
                    new CellRangeAddress(0, 0, 0, 6));

            XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(1, 1, 0, 6));

            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);



            //XDDFChartData.Series series1 = data.addSeries(countries, values);


            XDDFChartData.Series  series1 = data.addSeries(XDDFDataSourcesFactory.fromArray(cat)
                    , XDDFDataSourcesFactory.fromArray(val));



            series1.setTitle("Country", null);
            data.setVaryColors(true);
            chart.plot(data);

            // in order to transform a bar chart into a column chart, you just need to change the bar direction
            XDDFBarChartData bar = (XDDFBarChartData) data;
            //bar.setBarDirection(BarDirection.BAR);
            bar.setBarDirection(BarDirection.COL);

            // Write output to an excel file
            String filename = "files/bar-chart-top-seven-countries.xlsx";//"column-chart-top-seven-countries.xlsx";
            try (FileOutputStream fileOut = new FileOutputStream(filename)) {
                wb.write(fileOut);
            }
        }
    }

}
