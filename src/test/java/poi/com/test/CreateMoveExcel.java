package poi.com.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * http://blog.51cto.com/zangyanan/1837229 wang 2018/3/8
 */
public class CreateMoveExcel {

  public static void main(String[] args) throws Exception {
    FileInputStream is = new FileInputStream("F:\\data_doc\\历史数据导入\\课程表.xls");
    HSSFWorkbook workbook = new HSSFWorkbook(is);
    HSSFSheet sheet = workbook.getSheet("课程表");
    sheet.shiftRows(0, sheet.getLastRowNum(), 1);
    sheet.shiftRows(6, sheet.getLastRowNum(), 1);

    /**
     * 开始我认为移动会自己创建行和列，因此我直接
     * 用方法想获取row以及cell,这时候报空指针，查API了解
     * shiftRows可以把某区域的行移动，但是移动后剩下的区域却为空
     * 因此我们需要创建
     */
    /*
    HSSFRow row = sheet.getRow(0);
    HSSFCell cell = row.getCell(0);
    cell.setCellValue("课程表");
    */

    HSSFRow row = sheet.createRow(0);
    HSSFCell cell = row.createCell(0);
    cell.setCellValue("课程表");

    HSSFRow srow = sheet.createRow(6);
    HSSFCell scell= srow.createCell(0);
    scell.setCellValue("午休");

    /**
     * 合并单元格功能，对新增的第一行进行合并
     */
    CellRangeAddress address = new CellRangeAddress(0, 0, 0, 4);
    sheet.addMergedRegion(address);

    /**
     * 对新半的第六行进行合并
     */
    CellRangeAddress secondAddress = new CellRangeAddress(6, 6, 0, 4);
    sheet.addMergedRegion(secondAddress);

    /**
     * 对表格的修改以及其他 操作需要在workbook.write之后生效的
     */
    FileOutputStream os = new FileOutputStream("F:\\data_doc\\历史数据导入\\课程表.xls");
    workbook.write(os);
    is.close();
    os.close();


  }

}
