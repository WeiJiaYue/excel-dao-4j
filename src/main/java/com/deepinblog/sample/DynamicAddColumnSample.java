package com.deepinblog.sample;

import com.deepinblog.excel.ExcelProcessor;
import com.deepinblog.excel.ExcelTable;

import java.util.Map;

/**
 * Created by louisyuu
 */
public class DynamicAddColumnSample {


    /**
     * Project path
     */
    static String PROJECT_PATH = System.getProperty("user.dir");

    static String SRC_PATH = "/src/main/java/com/deepinblog/sample/";


    public static void main(String[] args) throws Exception {
        //Load Excel Template
        ExcelProcessor processor = new ExcelProcessor(PROJECT_PATH + SRC_PATH, "InsertToExcelSample.xls") {
            //Real process excel
            //You can do excel CRUD here
            @Override
            public void doProcess(ExcelTable table) throws Exception {
                //Add new column
                table.addColumn("MA5");
                table.addColumn("Profit");
                for (Map<String, Object> row : table.getRows()) {
                    //Set new column val
                    table.updateRow(row, "MA5", "This is ma5");
                    table.updateRow(row, "Profit", "This is Profit");
                }
            }
        };
        processor.setNeededGenerateNewExcel(true);
        processor.setNewFileName("DynamicAddColumnSample");
        //Generate new excel file
        processor.process();

    }


}
