package com.deepinblog.sample;

import com.deepinblog.excel.ExcelProcessor;
import com.deepinblog.excel.ExcelTable;

import java.util.List;
import java.util.Map;

/**
 * Created by louisyuu
 */
public class SelectSample {
    /**
     * Project path
     */
    static String PROJECT_PATH = System.getProperty("user.dir");

    static String SRC_PATH = "/src/main/java/com/deepinblog/sample/";


    public static void main(String[] args) throws Exception {
        //Load Excel Template
        ExcelProcessor processor = new ExcelProcessor(PROJECT_PATH + SRC_PATH, "SelectSample.xls") {
            //Real process excel
            //You can do excel CRUD here
            @Override
            public void doProcess(ExcelTable table) throws Exception {
                //Get rows by condition
                List<Map<String, Object>> rowsLikeColumn = table.selectRowsLikeColumn("Hello", "你好");
                List<Map<String, Object>> rowsEqualsColumn = table.selectRowsEqualsColumn("Test", "ThisTest");
                print("rowsEqualsColumn ", rowsEqualsColumn);
                System.out.println("=================");
                print("rowsLikeColumn ", rowsLikeColumn);
            }
        };
        processor.setNeededGenerateNewExcel(false);
        //Just run
        processor.process();
    }

    private static void print(String title, List<Map<String, Object>> rows) {
        for (Map row : rows) {
            System.out.println(title + ": " + row);
        }
    }
}
