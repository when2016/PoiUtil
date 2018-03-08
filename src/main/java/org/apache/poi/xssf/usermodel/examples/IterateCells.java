package org.apache.poi.xssf.usermodel.examples;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * wang 2018/3/8
 */
public class IterateCells {

  private static final String filePath = "F:\\data_doc\\历史数据导入\\精武门团体赛小组赛 2018.1.20 叶斯泰VS阿依肯.xlsx";

  public static void main(String[] args) throws IOException {

    args = new String[]{filePath};
    try (Workbook wb = new XSSFWorkbook(new FileInputStream(args[0]))) {
      for (int i = 3; i < wb.getNumberOfSheets(); i++) {
        Sheet sheet = wb.getSheetAt(i);
        System.out.println("sheetName" + wb.getSheetName(i));
        for (Row row : sheet) {
          System.out.println("rownum: " + row.getRowNum());
          for (Cell cell : row) {
            System.out.print(cell + ",");
          }
          System.out.println("");
        }
      }
    }
  }

}
