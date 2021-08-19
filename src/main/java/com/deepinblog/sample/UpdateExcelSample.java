package com.deepinblog.sample;

import com.deepinblog.excel.ExcelProcessor;
import com.deepinblog.excel.ExcelTable;

import java.util.List;
import java.util.Map;

/**
 * Created by louisyuu
 */
public class UpdateExcelSample {


    /**
     * Project path
     */
    static String PROJECT_PATH = System.getProperty("user.dir");

    static String SRC_PATH = "/src/main/java/com/deepinblog/sample/";


    public static void main(String[] args) throws Exception {
        //Excel to be updated
        ExcelProcessor processor = new ExcelProcessor(PROJECT_PATH + SRC_PATH, "InsertToExcelSample.xls") {
            //Real process excel
            //You can do excel CRUD here
            @Override
            public void doProcess(ExcelTable table) throws Exception {
                //Update excel
                for (Map<String, Object> row : table.getRows()) {
                    table.updateRow(row, "Low", "OverrideAll");
                }
            }
        };
        processor.setNeededGenerateNewExcel(true);
        processor.setNewFileName("UpdateExcelSample");
        //Generate new excel file
        processor.process();

    }


}
