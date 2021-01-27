package com.mcubes.writer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.List;

public class ExcelWriter {

    private enum DataType {
        BOOLEAN, STRING, INTEGER, FLOAT, DATE
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
    private List<Float[]> floatData;
    private List<Date[]> dateData;
    private DataType selectedType;

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

    public void setFloatData(List<Float[]> floatData) {
        this.floatData = floatData;
        this.selectedType = DataType.FLOAT;
    }

    public void setDateData(List<Date[]> dateData) {
        this.dateData = dateData;
        this.selectedType = DataType.DATE;
    }

    public void write() throws IOException {
        this.createHeadersRow();
        this.setData();
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
                    case FLOAT:
                        for (int i = 0; i < floatData.size(); i++) {
                            row = sheet.createRow(i + 1);
                            Float[] dataRow = floatData.get(i);
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

}
