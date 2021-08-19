package com.deepinblog.sample;

import com.deepinblog.excel.ExcelProcessor;
import com.deepinblog.excel.ExcelTable;

import java.util.*;

/**
 * Created by louisyuu
 */
public class InsertToExcelSample {
    /**
     * Project path
     */
    static String PROJECT_PATH = System.getProperty("user.dir");

    static String SRC_PATH = "/src/main/java/com/deepinblog/sample/";


    public static void main(String[] args) throws Exception {
        //Load excel template as your new excel data structure
        ExcelProcessor processor = new ExcelProcessor(PROJECT_PATH + SRC_PATH, "KlineTemplate.xlsx") {
            //Real process excel
            //You can do excel CRUD here
            @Override
            public void doProcess(ExcelTable table) throws Exception {
                //You can initial your excel data from API,Database,whatever.as your wish
                //Mock inserts five rows
                for (int i = 0; i < 5; i++) {
                    Map<String, Object> emptyRow = table.createEmptyRow();
                    emptyRow.put("Timestamp", System.currentTimeMillis());
                    emptyRow.put("Open", 1000 + i);
                    emptyRow.put("High", 2000 + i);
                    emptyRow.put("Low", 3000 + i);
                    emptyRow.put("Close", 4000 + i);
                    emptyRow.put("Volume", 5000 + i);
                    table.addRow(emptyRow);
                }
            }
        };
        processor.setNeededGenerateNewExcel(true);
        processor.setNewFileName("InsertToExcelSample");
        //Generate new excel file
        processor.process();
    }
}
