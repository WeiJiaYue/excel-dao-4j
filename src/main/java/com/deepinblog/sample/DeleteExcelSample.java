package com.deepinblog.sample;

import com.deepinblog.excel.ExcelProcessor;
import com.deepinblog.excel.ExcelTable;

import java.util.Map;

/**
 * Created by louisyuu
 */
public class DeleteExcelSample {


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
                if(table.deleteRow(4)){
                    System.out.println("Row index 4 has been deleted");
                }else{
                    System.out.println("Row index out of bound");
                }
            }
        };
        processor.setNeededGenerateNewExcel(true);
        processor.setNewFileName("DeleteExcelSample");
        //Generate new excel file
        processor.process();
    }
}
