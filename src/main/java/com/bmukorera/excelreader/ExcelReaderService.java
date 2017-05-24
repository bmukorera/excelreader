package com.bmukorera.excelreader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

/**
 * Created by bmukorera on 5/16/2017.
 *
 */

public class ExcelReaderService {
    private int numberOfSheets;
    public Map<Long,List<ExcelData>> convertExcelDocument(FileInputStream myExcelFile,String dataDeliminitor){
        Map<Long,List<ExcelData>> response = new HashMap<>();
        try {
            Workbook workbook = new XSSFWorkbook(myExcelFile);
            //assign number of sheets
            numberOfSheets = workbook.getNumberOfSheets();
            for(int sheetCounter=1;sheetCounter<=numberOfSheets;sheetCounter++){
                Sheet dataSheetToBeRead = workbook.getSheetAt(sheetCounter-1);
                int rowCounter=1;
                Iterator<Row> iterator = dataSheetToBeRead.iterator();
                List<ExcelData> listOfExcelRowData=new ArrayList<>();
                while (iterator.hasNext()) {
                    Row currentRow=iterator.next();
                    StringBuilder excelDataforRow=new StringBuilder("");
                    for(int cellNum=0;cellNum<currentRow.getLastCellNum();cellNum++){
                        Cell currentCell = currentRow.getCell(cellNum,Row.CREATE_NULL_AS_BLANK);
                        if (currentCell.getCellTypeEnum() == CellType.STRING) {
                            excelDataforRow.append(currentCell.getStringCellValue());
                        } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                            excelDataforRow.append(currentCell.getNumericCellValue());
                        }
                        else if(currentCell.getCellTypeEnum()==CellType.BLANK){
                            excelDataforRow.append("");
                        }
                        else if(currentCell.getCellTypeEnum()==CellType._NONE){
                            excelDataforRow.append("");
                        }
                        else if(currentCell.getCellTypeEnum()==CellType.BOOLEAN){
                            excelDataforRow.append(currentCell.getBooleanCellValue());
                        }
                        else if(currentCell.getCellTypeEnum()==CellType.ERROR){
                            excelDataforRow.append(currentCell.getErrorCellValue());
                        }
                        else if(currentCell.getCellTypeEnum()==CellType.FORMULA){
                            excelDataforRow.append(currentCell.getStringCellValue());
                        }else{
                            excelDataforRow.append("");
                        }
                        excelDataforRow.append(dataDeliminitor);
                    }
                    listOfExcelRowData.add(new ExcelData(rowCounter,excelDataforRow.toString(),currentRow.getLastCellNum()));
                    rowCounter++;
                }
                response.put(new Long(sheetCounter),listOfExcelRowData);
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
return response;
    }
}
