package poi.com.test;

import java.io.FileOutputStream;
import java.util.Random;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * http://blog.51cto.com/zangyanan/1837229
 * wang 2018/3/7
 */
public class CreateSimpleExcel {

  public static void main(String[] args) throws Exception {
    HSSFWorkbook workbook = new HSSFWorkbook();
    HSSFSheet sheet = workbook.createSheet("课程表");

    HSSFRow row = sheet.createRow(0);
    row.createCell(0).setCellValue("星期一");
    row.createCell(1).setCellValue("星期二");
    row.createCell(2).setCellValue("星期三");
    row.createCell(3).setCellValue("星期四");
    row.createCell(4).setCellValue("星期五");

    Random random = new Random();
    String[] course = {"语文", "数学", "英语", "物理", "化学", "政治", "历史", "音乐", "美术", "体育"};
    for (int j = 1; j <= 7; j++) {
      HSSFRow row_j = sheet.createRow(j);
      for (int k = 0; k <= 4; k++) {
        int i = random.nextInt(10);
        System.out.println(i);
        row_j.createCell(k).setCellValue(course[i]);
      }
    }

    FileOutputStream out = new FileOutputStream("F:\\data_doc\\历史数据导入\\课程表.xls");
    workbook.write(out);
    out.close();;


  }

}
