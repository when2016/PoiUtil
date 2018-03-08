package poi.com.test;

import java.io.FileOutputStream;
import java.util.Random;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * http://blog.51cto.com/zangyanan/1837229 wang 2018/3/8
 */
public class createSimple2Excel {

  public static void main(String[] args) throws Exception {
    HSSFWorkbook workbook = new HSSFWorkbook();
    /**
     * createSheet存在有参和无参两种形式，主要设置sheet名字
     */
    HSSFSheet sheet = workbook.createSheet("课程表");
    HSSFRow row = sheet.createRow(0);
    HSSFCell cell = row.createCell(0);
    cell.setCellValue("课程表");
    CellRangeAddress address = new CellRangeAddress(0, 0, 0, 5);
    sheet.addMergedRegion(address);
    /**
     * 我们知道课程表第一行是代表周一到周五,下面我们用两种方式创建Cell,
     * 一种用变量，另一种未用变量,用变量的好处后面可以体会到。
     */
    HSSFRow secondRow = sheet.createRow(1);//创建第二行
    //这里面我们第一列不用是因为第三行存在合并的上午单元格，自己体会下
    secondRow.createCell(1).setCellValue("星期一");
    secondRow.createCell(2).setCellValue("星期二");
    secondRow.createCell(3).setCellValue("星期三");
    secondRow.createCell(4).setCellValue("星期四");
    secondRow.createCell(5).setCellValue("星期五");

    /**
     * 上面我们只是设置了首行，后面课程我们用同样的方法设置，
     * 这里面我们用循环方法设置课程
     */
    Random random = new Random();
    String[] course = {"语文", "数学", "英语", "物理", "化学", "政治", "历史", "音乐", "美术", "体育"};
    //循环产生7个row;
    for (int j = 2; j <= 9; j++) {
      //每个row的1-5个cell设置值,我用随机取数组来写值。
      HSSFRow row_j = sheet.createRow(j);
      //第六行是午休
      if (j == 6) {
        row_j.createCell(0).setCellValue("午休");
        CellRangeAddress secondAddress = new CellRangeAddress(6, 6, 0, 5);
        sheet.addMergedRegion(secondAddress);
        continue;
      }
      //每行开始都要空出一列来让我们能增加上午下午单元格
      for (int k = 1; k <= 5; k++) {
        int i = random.nextInt(10);
        row_j.createCell(k).setCellValue(course[i]);
      }
    }

    sheet.getRow(2).createCell(0).setCellValue("上午");
    sheet.getRow(7).createCell(0).setCellValue("下午");
    CellRangeAddress thirdAddress = new CellRangeAddress(2, 5, 0, 0);
    sheet.addMergedRegion(thirdAddress);

    CellRangeAddress fourthAddress = new CellRangeAddress(7, 9, 0, 0);
    sheet.addMergedRegion(fourthAddress);

    FileOutputStream out = new FileOutputStream("F:\\data_doc\\历史数据导入\\课程表.xls");
    workbook.write(out);
    out.close();

  }


}
