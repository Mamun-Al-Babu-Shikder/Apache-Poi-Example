package com.mcubes;

import com.mcubes.model.CellData;
import com.mcubes.model.CellValue;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class ProjectEffectivenessReportsUsingTemplate {

    public static void main(String[] args) {

        List<CellValue> dataList = new ArrayList<>();
        dataList.add(new CellValue(4, 3, "WORKSPACE NAME"));
        dataList.add(new CellValue(5, 3, "PROJECT NAME"));
        dataList.add(new CellValue(6, 3, Integer.toString(849000)));
        dataList.add(new CellValue(7, 3, Integer.toString(80)));
        dataList.add(new CellValue(8, 3, Integer.toString(100)));

        dataList.add(new CellValue(12, 3, Integer.toString(300000)));
        dataList.add(new CellValue(13, 3, Integer.toString(212000)));
        dataList.add(new CellValue(14, 3, Integer.toString(310000)));
        dataList.add(new CellValue(15, 3, Integer.toString(215000)));


        dataList.add(new CellValue(19, 3, Integer.toString(385)));
        dataList.add(new CellValue(20, 3, Integer.toString(20)));

        dataList.add(new CellValue(23, 5, Double.toString(3.20)));
        dataList.add(new CellValue(23, 9, Double.toString(7.91)));

        createReport(dataList);
    }

    public static void createReport(List<CellValue> data) {

        try {
            File file = new File("files/Project_Effectiveness_Reports_CALl.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(1);

            for (CellValue info : data) {
                Row row = sheet.getRow(info.getRow() -  1);
                Cell cell = row.getCell(info.getColumn() - 1);
                cell = cell == null ? row.createCell(info.getColumn() - 1) : cell;
                cell.setCellValue(info.getValue());
                System.out.println(info);
            }

            FileOutputStream outputStream = new FileOutputStream("files/Report_cal_" + System.currentTimeMillis() +
                    ".xlsx");
            workbook.write(outputStream);
            workbook.close();

        } catch (Exception e){
            e.printStackTrace();
        }

    }

//    public static File copyExcelTemplate() {
//        try {
//            File source = new File("files/Project_Effectiveness_Reports_CALl.xlsx");
//            File dest = new File("files/Report_cal_" + System.currentTimeMillis() + ".xlsx");
//            FileUtils.copyFile(source, dest);
//            return dest;
//        }catch (Exception e){
//            e.printStackTrace();
//            return null;
//        }
//    }

}
