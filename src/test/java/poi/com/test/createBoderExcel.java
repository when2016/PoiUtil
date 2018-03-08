package poi.com.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * wang 2018/3/8
 */
public class createBoderExcel {

  private static final String filePath = "F:\\data_doc\\历史数据导入\\课程表.xls";

  public static void main(String[] args) {
    try {
      FileInputStream is = new FileInputStream(filePath);
      HSSFWorkbook workbook = new HSSFWorkbook(is);
      HSSFSheet sheet = workbook.getSheet("课程表");
      HSSFCellStyle firststyle = workbook.createCellStyle();//第一种样式针对第一个单元格的，不存在右边线
//      firststyle.setBorderTop(HSSFCellStyle.BORDER_THICK);
//      firststyle.setBorderLeft(HSSFCellStyle.BORDER_THICK);
//      firststyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
//      firststyle.setTopBorderColor(HSSFColor.PINK.index);
//      firststyle.setLeftBorderColor(HSSFColor.PINK.index);
//      firststyle.setBottomBorderColor(HSSFColor.BLUE.index);

      HSSFCellStyle secondstyle = workbook.createCellStyle();//第二种样式针对中间单元格的，不存在左右边线
//      secondstyle.setBorderTop(HSSFCellStyle.BORDER_THICK);
//      secondstyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
      secondstyle.setTopBorderColor(HSSFColor.PINK.index);
      secondstyle.setBottomBorderColor(HSSFColor.BLUE.index);

      HSSFCellStyle thirdstyle = workbook.createCellStyle();//第三种样式针对最后单元格的，不存在左边线
//      thirdstyle.setBorderTop(HSSFCellStyle.BORDER_THICK);
//      thirdstyle.setBorderRight(HSSFCellStyle.BORDER_THICK);
//      thirdstyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
      thirdstyle.setTopBorderColor(HSSFColor.PINK.index);
      thirdstyle.setRightBorderColor(HSSFColor.PINK.index);
      thirdstyle.setBottomBorderColor(HSSFColor.BLUE.index);
      HSSFRow firstrow = sheet.getRow(0);
      for (int i = 0; i < firstrow.getLastCellNum(); i++) {
        if (i == 0) {
//          firststyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
          firstrow.getCell(i).setCellStyle(firststyle);
        } else if (i == firstrow.getLastCellNum() - 1) {
          firstrow.createCell(i);//注意前面实例针对第一行只创建了第一列，因此在这里我们需要创建列，不然不会设置边框
          firstrow.getCell(i).setCellStyle(thirdstyle);
        } else {
          firstrow.createCell(i);
          firstrow.getCell(i).setCellStyle(secondstyle);
        }
      }
      //发现对合并的单元格设置边框，居中效果居然没了，因此我们在这里补充
      FileOutputStream out = new FileOutputStream(filePath);
      workbook.write(out);
      out.close();
      is.close();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}
