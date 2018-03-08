package poi.com.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * http://blog.51cto.com/zangyanan/1837229 wang 2018/3/8
 */
public class CreateStyleExcel {

  private static final String filePath = "F:\\data_doc\\历史数据导入\\课程表.xls";

  public static void main(String[] args) {
    try {
      FileInputStream is = new FileInputStream(filePath);
      HSSFWorkbook workbook = new HSSFWorkbook(is);
      HSSFSheet sheet = workbook.getSheet("课程表");
      HSSFRow firstRow = sheet.getRow(0);//获取课程表行
      HSSFRow secondRow = sheet.getRow(2);//获取上午行
      HSSFRow sixRow = sheet.getRow(6);//获取午休行
      HSSFRow sevenRow = sheet.getRow(7);//获取下午行

      HSSFCellStyle style = workbook.createCellStyle();
      style.setAlignment(HorizontalAlignment.CENTER);//水平居中
      style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
      firstRow.getCell(0).setCellStyle(style);
      secondRow.getCell(0).setCellStyle(style);
      sixRow.getCell(0).setCellStyle(style);
      sevenRow.getCell(0).setCellStyle(style);

      FileOutputStream out = new FileOutputStream(filePath);
      workbook.write(out);

      out.close();
      is.close();


    } catch (Exception e) {
      e.printStackTrace();
    }
  }

}
