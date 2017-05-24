package com.bmukorera.excelreader;

/**
 * Created by bmukorera on 5/16/2017.
 */
public class ExcelData {
    long row;
    String rowContent;
    long numberOfColumns;

    public ExcelData(long row, String rowContent,long numberOfColumns) {
        this.row = row;
        this.rowContent = rowContent;
        this.numberOfColumns=numberOfColumns;
    }

    public ExcelData() {
    }

    public long getRow() {
        return row;
    }

    public void setRow(long row) {
        this.row = row;
    }

    public String getRowContent() {
        return rowContent;
    }

    public void setRowContent(String rowContent) {
        this.rowContent = rowContent;
    }

    public long getNumberOfColumns() {
        return numberOfColumns;
    }

    public void setNumberOfColumns(long numberOfColumns) {
        this.numberOfColumns = numberOfColumns;
    }

    @Override
    public String toString() {
        return "\nExcelData{" +
                "row=" + row +
                ", rowContent='" + rowContent + '\'' +
                ", numberOfColumns=" + numberOfColumns +
                '}';
    }
}
