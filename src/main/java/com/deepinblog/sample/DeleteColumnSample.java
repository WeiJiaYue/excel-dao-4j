package com.deepinblog.sample;

import com.deepinblog.excel.ExcelProcessor;
import com.deepinblog.excel.ExcelTable;

import java.util.Map;

/**
 * Created by louisyuu
 */
public class DeleteColumnSample {


    /**
     * Project path
     */
    static String PROJECT_PATH = System.getProperty("user.dir");

    static String SRC_PATH = "/src/main/java/com/deepinblog/sample/";


    public static void main(String[] args) throws Exception {
        //Load Excel Template
        ExcelProcessor processor = new ExcelProcessor(PROJECT_PATH + SRC_PATH, "DynamicAddColumnSample.xls") {
            //Real process excel
            //You can do excel CRUD here
            @Override
            public void doProcess(ExcelTable table) throws Exception {
                //Add new column
                table.deleteColumn("Profit");
                table.deleteColumn("Low");
            }
        };
        processor.setNeededGenerateNewExcel(true);
        processor.setNewFileName("DeleteColumnSample");
        //Generate new excel file
        processor.process();
    }


}
