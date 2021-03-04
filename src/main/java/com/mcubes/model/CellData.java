package com.mcubes.model;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class CellData {

    public static enum DataType {
        BLANK, BOOLEAN, DOUBLE, STRING, CHART
    }

    private int row;
    private int column;
    private int rowspan;
    private int colspan;
    private boolean bold;
    private HorizontalAlignment horizontalAlignment;
    private VerticalAlignment verticalAlignment;
    private String stringValue;
    private boolean booleanValue;
    private double doubleValue;
    private DataType dataType = DataType.BLANK;
    private String formula;
    private String format;

    public CellData(int row, int column) {
        this.row = row;
        this.column = column;
    }

    public CellData(int row, int column, boolean value) {
        this(row, column);
        this.booleanValue = value;
        dataType = DataType.BOOLEAN;
    }

    public CellData(int row, int column, double value) {
        this(row, column);
        this.doubleValue = value;
        dataType = DataType.DOUBLE;
    }

    public CellData(int row, int column, String value) {
        this(row, column);
        this.stringValue = value;
        dataType = DataType.STRING;
    }

    public CellData(int row, int column, String formula, String format) {
        this(row, column);
        this.formula = formula;
        this.format = format;
    }

    public CellData(int row, int column, double value, String formula, String format) {
        this(row, column, formula, format);
        this.doubleValue = value;
        dataType = DataType.DOUBLE;
    }

    public CellData(int row, int column, String formula, String format, HorizontalAlignment horizontalAlignment, int colspan) {
        this(row, column, formula, format);
        this.horizontalAlignment = horizontalAlignment;
        this.colspan = colspan;
    }

    public CellData(int row, int column, String value, boolean boldText) {
        this(row, column);
        this.stringValue = value;
        dataType = DataType.STRING;
        this.bold = boldText;
    }

    public CellData(int row, int column, double value, boolean boldText) {
        this(row, column);
        this.doubleValue = value;
        dataType = DataType.DOUBLE;
        this.bold = boldText;
    }

    public CellData(int row, int column, String value, boolean boldText, HorizontalAlignment alignment) {
        this(row, column, value);
        dataType = DataType.STRING;
        this.bold = boldText;
        this.horizontalAlignment = alignment;
    }

    public CellData(int row, int column, String value, boolean boldText, HorizontalAlignment alignment, int colspan) {
        this(row, column, value, boldText, alignment);
        this.colspan = colspan;
    }

    public int getRow() {
        return row - 1;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getColumn() {
        return column - 1;
    }

    public void setColumn(int column) {
        this.column = column;
    }

    public int getRowspan() {
        return rowspan;
    }

    public void setRowspan(int rowspan) {
        this.rowspan = rowspan;
    }

    public int getColspan() {
        return colspan;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }

    public boolean isBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public void setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public DataType getDataType() {
        return dataType;
    }

    public void setValue(boolean value) {
        this.booleanValue = value;
        this.dataType = DataType.BOOLEAN;
    }

    public void setValue(double value) {
        this.doubleValue = value;
        this.dataType = DataType.DOUBLE;
    }

    public void setValue(String value) {
        this.stringValue = value;
        this.dataType = DataType.STRING;
    }

    public String getStringValue() {
        return stringValue;
    }

    public boolean getBooleanValue() {
        return booleanValue;
    }

    public double getDoubleValue() {
        return doubleValue;
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }
}
