


## Using Java to manipulate Excel's addition, deletion, modification, and query in batches is as simple as operating a Java collection



[excel-dao-4j](https://github.com/louis-yuu/excel-dao-4j)

# Quick start



### Sample 1,Add data to empty excel

* According to the data structure determined by the Excel template,In doProcess method you can get data from API, Database and other places for filling. Then generate the data into a new Excel file

* Before process

  ![image-20210819145710120](/imgs/sample-1-pre.png)


* After proceed

  ​	![image-20210819145335055](/imgs/sample-1-post.png)

* [Sample code](https://github.com/louis-yuu/excel-dao-4j/blob/master/src/main/java/com/deepinblog/sample/InsertToExcelSample.java)

  ```java
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
  
  ```




### Sample 2,Modify data

* Before process

  ![image-20210819145710120](/imgs/sample-2-pre.png)

* After proceed

  ![image-20210819145857082](/imgs/sample-2-post.png)

* [Sample code](https://github.com/louis-yuu/excel-dao-4j/blob/master/src/main/java/com/deepinblog/sample/UpdateExcelSample.java)

  ```java
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
  ```



### Sample 3,Delete excel row

* Before process

  ![image-20210819145710120](/imgs/sample-3-pre.png)

* After proceed

  ![image-20210819150513884](/imgs/sample-3-post.png)

* [Sample code](https://github.com/louis-yuu/excel-dao-4j/blob/master/src/main/java/com/deepinblog/sample/DeleteExcelSample.java)

  ```java
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
  
  ```



### Sample 4，Dynamic add new columns

* Before process

  ![image-20210819145710120](/imgs/sample-4-pre.png)

* After proceed

  ![image-20210819151102688](/imgs/sample-4-post.png)

* [Sample code](https://github.com/louis-yuu/excel-dao-4j/blob/master/src/main/java/com/deepinblog/sample/DynamicAddColumnSample.java)

  ```java
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
  ```



### Sample 5,Delete columns

* Before process

  ![image-20210819151328856](/imgs/sample-5-pre.png)

* After proceed

  ![image-20210819151452410](/imgs/sample-5-post.png)

* [Sample code](https://github.com/louis-yuu/excel-dao-4j/blob/master/src/main/java/com/deepinblog/sample/DeleteColumnSample.java)

  ```java
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
  ```



### Sample 6,Selection,According to column name Equivalent query, Like query

* Excel file to be queried

  ![image-20210819153223256](/imgs/sample-6-pre.png)

* Query results

  ![image-20210819153728391](/imgs/sample-6-post.png)

* [Sample code](https://github.com/louis-yuu/excel-dao-4j/blob/master/src/main/java/com/deepinblog/sample/SelectSample.java)

  ```java
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
  ```

## More powerful functions ,pls stay tuned

