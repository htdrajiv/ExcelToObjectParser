package com.javalearning.library;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.reflect.TypeToken;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class ExcelFileReader {
    private Logger logger = LogManager.getLogger(ExcelFileReader.class);

    /*
    * @param filePath: excel file path which contains information that we are looking for our type T objects.
    * @param sheetName: name of sheet in @filePath excel file which contains information for our type T objects.
    * @param clazz: target class type into which we are mapping data from excel file.
    * @return list of T type objects
    * */
    public <T> List<T> parse(String filePath, String sheetName, Class<T> clazz){
        Gson gson = new Gson();
        JsonObject jsonObject = getExcelDataAsJsonObject(new File(filePath), sheetName);
        logger.info("Now converting json object into "+clazz.getSimpleName()+" object...");
        List<T> results = gson.fromJson(jsonObject.get(sheetName).toString(),TypeToken.getParameterized(List.class, clazz).getType());
        logger.info("Done converting json object into "+clazz.getSimpleName()+" object.\n");
        return results;
    }

    private JsonObject getExcelDataAsJsonObject(File excelFile, String sheetName) {
        logger.info("Started reading excel file "+ excelFile.getName()+"...");
        JsonObject sheetsJsonObject = new JsonObject();
        Workbook workbook = null;
        try {
            workbook = new XSSFWorkbook(excelFile);
            JsonArray sheetArray = new JsonArray();
            ArrayList<String> columnNames = new ArrayList<String>();
            List<String> sheetNames = new ArrayList<String>();
            for (int i=0; i<workbook.getNumberOfSheets(); i++) {
                sheetNames.add( workbook.getSheetName(i) );
            }
            Sheet sheet = workbook.getSheet(sheetName);
            DataFormatter dataFormatter = new DataFormatter();
            for (Row currentRow : sheet) {
                JsonObject jsonObject = new JsonObject();
                if (currentRow.getRowNum() != 0) {
                    for (int j = 0; j < columnNames.size(); j++) {
                        if (currentRow.getCell(j) != null) {
                            if (currentRow.getCell(j).getCellType() == CellType.BLANK) {
                                jsonObject.addProperty(columnNames.get(j), "");
                            }else{
                                jsonObject.addProperty(columnNames.get(j), dataFormatter.formatCellValue(currentRow.getCell(j)));
                            }
                        } else {
                            jsonObject.addProperty(columnNames.get(j), "");
                        }
                    }
                    sheetArray.add(jsonObject);
                } else {
                    // if first row, then column names
                    for (int k = 0; k < currentRow.getPhysicalNumberOfCells(); k++) {
                        columnNames.add(currentRow.getCell(k).getStringCellValue());
                    }
                }
            }
            sheetsJsonObject.add(sheet.getSheetName(), sheetArray);
            System.out.println("sheetsJsonObject = " + sheetsJsonObject.toString());

            logger.info("Done reading and converting "+sheet.getSheetName()+" sheet from excel file into json object.");
        } catch (Exception e) {
            logger.error("ExcelUtils -> getExcelDataAsJsonObject() :: Exception thrown constructing XSSFWorkbook from provided excel file.  InvalidFormatException | IOException => ");
            e.printStackTrace();
        }
        return sheetsJsonObject;
    }
}
