哈囉
這是一個讀取 xlsx 的程式

執行方式：

安裝 maven

在 test 目錄下執行 mvn clean install，產生 target/test-0.0.1-SNAPSHOT.jar

將 target/test-0.0.1-SNAPSHOT.jar 與 src/main/resources/Q1.xlsx 放在同一個目錄下

在 test-0.0.1-SNAPSHOT.jar 與 Q1.xlsx 放置的目錄下執行終端機，輸入 java -jar test-0.0.1-SNAPSHOT.jar

按下 enter，即可獲取結果

注意：

只會讀取檔名為 Q1.xlsx 的 第一個工作表
