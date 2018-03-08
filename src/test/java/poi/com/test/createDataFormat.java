package poi.com.test;

import java.io.FileOutputStream;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * wang 2018/3/8
 */
public class createDataFormat {

  private static final String filePath = "F:\\data_doc\\历史数据导入\\格式转换.xls";

  public static void main(String[] args) {
    try {
      HSSFWorkbook workbook = new HSSFWorkbook();
      HSSFSheet sheet = workbook.createSheet("格式转换");
      HSSFRow row0 = sheet.createRow(0);
      /**
       * 时间格式转换
       * 我们用第一排第一个、第二个、第三个单元格都设置当前时间
       * 然后第一个单元格不进行任何操作，第二个单元格用内嵌格式，第三个单元格用自定义
       */
      Date date = new Date();
      HSSFCell row1_cell1 = row0.createCell(0);
      HSSFCell row1_cell2 = row0.createCell(1);
      HSSFCell row1_cell3 = row0.createCell(2);
      row1_cell1.setCellValue(date);
      row1_cell2.setCellValue(date);
      row1_cell3.setCellValue(date);
      HSSFCellStyle style1 = workbook.createCellStyle();
      style1.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
      HSSFCellStyle style2 = workbook.createCellStyle();
      style2.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd hh:m:ss"));
      row1_cell2.setCellStyle(style1);
      row1_cell3.setCellStyle(style2);
      /**
       * 第二排我们进行小数处理
       * 第一个不进行任何处理，第二个我们用内嵌格式保留两位，第三个我们用自定义
       */
      HSSFRow row1 = sheet.createRow(1);
      double db = 3.1415926;
      HSSFCell row2_cell1 = row1.createCell(0);
      HSSFCell row2_cell2 = row1.createCell(1);
      HSSFCell row2_cell3 = row1.createCell(2);
      row2_cell1.setCellValue(db);
      row2_cell2.setCellValue(db);
      row2_cell3.setCellValue(db);
      HSSFCellStyle style3 = workbook.createCellStyle();
      style3.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
      HSSFCellStyle style4 = workbook.createCellStyle();
      style4.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
      row2_cell2.setCellStyle(style3);
      row2_cell3.setCellStyle(style4);
      /**
       * 下面是进行货币的三种形式
       */
      HSSFRow row2 = sheet.createRow(2);
      double money = 12345.6789;
      HSSFCell row3_cell1 = row2.createCell(0);
      HSSFCell row3_cell2 = row2.createCell(1);
      HSSFCell row3_cell3 = row2.createCell(2);
      row3_cell1.setCellValue(money);
      row3_cell2.setCellValue(money);
      row3_cell3.setCellValue(money);
      HSSFCellStyle style5 = workbook.createCellStyle();
      style5.setDataFormat(HSSFDataFormat.getBuiltinFormat("￥#,##0.00"));
      HSSFCellStyle style6 = workbook.createCellStyle();
      style6.setDataFormat(workbook.createDataFormat().getFormat("￥#,##0.00"));
      row3_cell2.setCellStyle(style3);
      row3_cell3.setCellStyle(style4);
      FileOutputStream out = new FileOutputStream(filePath);
      workbook.write(out);
      out.close();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

}
