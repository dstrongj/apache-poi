// must have in CLASSPATH
// ojdbc6.jar; “C:\poi-3.9\poi-3.9-20121203.jar;” “C:\poi-3.9\poi-ooxml-3.9-20121203.jar;” “C:\poi-3.9\poi-ooxml-schemas-3.9-20121203.jar;” “C:\poi-3.9\ooxml-lib\dom4j-1.6.1.jar;”
// “C:\poi-3.9\ooxml-lib\xmlbeans-2.3.0.jar;


import java.io.File;
import java.io.FileOutputStream;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDatabase {
   public static void main(String[] args) throws Exception {
      Class.forName("oracle.jdbc.driver.OracleDriver");
      Connection connect = DriverManager.getConnection("jdbc:oracle:thin:un/uwd@host:port:sid");//
      
      Statement statement = connect.createStatement();
      ResultSet resultSet = statement.executeQuery("select * from database");
      XSSFWorkbook workbook = new XSSFWorkbook(); 
      XSSFSheet spreadsheet = workbook.createSheet("sheet_name");
      
      XSSFRow row = spreadsheet.createRow(1);
      XSSFCell cell;
      cell = row.createCell(1);
      cell.setCellValue("NAME");
      cell = row.createCell(2);
      cell.setCellValue("ID");
      cell = row.createCell(3);
      cell.setCellValue("LID");
      cell = row.createCell(4);
      cell.setCellValue("RID");
      cell = row.createCell(5);
      cell.setCellValue("API");
      int i = 2;

      while(resultSet.next()) {
         row = spreadsheet.createRow(i);
         cell = row.createCell(1);
         cell.setCellValue(resultSet.getInt("NAME"));
         cell = row.createCell(2);
         cell.setCellValue(resultSet.getString("ID"));
         cell = row.createCell(3);
         cell.setCellValue(resultSet.getString("LID"));
         cell = row.createCell(4);
         cell.setCellValue(resultSet.getString("RID"));
         cell = row.createCell(5);
         cell.setCellValue(resultSet.getString("API"));
         i++;
      }

      FileOutputStream out = new FileOutputStream(new File("UNKNOWN.xlsx"));
      workbook.write(out);
      out.close();
      System.out.println("UNKNOWN.xlsx written successfully");
   }
}