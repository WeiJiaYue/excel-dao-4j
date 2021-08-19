


## 用Java批量操作Excel的增删改查就像操作Java集合一样简单



[excel-dao-4j](https://github.com/louis-yuu/excel-dao-4j)

# 使用示例



### 示例一，对一个空Excel批量新增数据

* 在doProcess中根据Excel模板确定好的数据结构，你可以从API，Database等任何地方获取数据进行填充。然后生成数据到一个新的Excel文件中

* 加工前的Excel文件
  
  ![image-20210819145710120](/imgs/sample-1-pre.png)


* 加工后新生成Excel文件

  ​	![image-20210819145335055](/Users/lewis/Library/Application Support/typora-user-images/image-20210819145335055.png)

* 代码示例

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



### 示例二，更新Excel中的数据

* 加工前的Excel文件

  ![image-20210819145710120](/Users/lewis/Library/Application Support/typora-user-images/image-20210819145710120.png)

* 加工后的Excel文件

  ![image-20210819145857082](/Users/lewis/Library/Application Support/typora-user-images/image-20210819145857082.png)

* 代码示例

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



### 示例三，删除Excel中行数据

* 加工前的Excel文件

  ![image-20210819145710120](/Users/lewis/Library/Application Support/typora-user-images/image-20210819145710120.png)

* 加工后的Excel文件

  ![image-20210819150513884](/Users/lewis/Library/Application Support/typora-user-images/image-20210819150513884.png)

* 代码示例

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



###  

### 示例四，动态新增列

* 加工前的Excel文件

  ![image-20210819145710120](/Users/lewis/Library/Application Support/typora-user-images/image-20210819145710120.png)

* 加工后的Excel文件

  ![image-20210819151102688](/Users/lewis/Library/Application Support/typora-user-images/image-20210819151102688.png)

* 代码示例

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



### 示例五，删除列

* 加工前的Excel文件

  ![image-20210819151328856](/Users/lewis/Library/Application Support/typora-user-images/image-20210819151328856.png)

* 加工后的Excel文件

  ![image-20210819151452410](/Users/lewis/Library/Application Support/typora-user-images/image-20210819151452410.png)

* 代码示例

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



### 示例六，根据列名等值查询，Like查询

* 查询Excel如图

  ![image-20210819153223256](/Users/lewis/Library/Application Support/typora-user-images/image-20210819153223256.png)

* 查询效果展示

  ![image-20210819153728391](/Users/lewis/Library/Application Support/typora-user-images/image-20210819153728391.png)

* 代码示例

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

