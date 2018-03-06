package com.wanghongen.utils;

import java.io.FileInputStream;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * 一个例子说明POI解析EXCEL的大致原理：（取自网上已有–一个粗糙例子）取自此文 某管理员要查某层楼有多少人叫什么名字？ 1)首先要明确大楼在那里(找到对应的文件)
 * 2)其次要明确是在第几单元(找到对应的sheet) 3)在找到第几层楼(对应的row) 4)敲门问住户户主先生/小姐的姓名(cell)
 */
public class TestA {

  public static void main(String args[]) throws Exception {
    //找到大楼的位置
    FileInputStream input = new FileInputStream("d:\\dir.xls");
    //告诉管理员
    POIFSFileSystem f = new POIFSFileSystem(input);
    //走到大楼楼下
    HSSFWorkbook wb = new HSSFWorkbook(f);
    //确认自己走到第几单元
    HSSFSheet sheet = wb.getSheetAt(0);
    //看一看有没有楼层
    Iterator rows = sheet.rowIterator();
    while (rows.hasNext()) {
      //如果有我们一层层问
      HSSFRow row = (HSSFRow) rows.next();
      Iterator cells = row.cellIterator();
      //如果有人开门
      while (cells.hasNext()) {
        //我们一户一户的登记
        HSSFCell cell = (HSSFCell) cells.next();
        //是先生还是小姐(对应的数据类型)
        int cellType = cell.getCellType();
        //System.out.print(getValue(cell,cellType));
      }
      System.out.println("");
    }
  }

}
