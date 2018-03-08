package com.wanghongen.utils;

import java.io.File;
import java.util.ArrayList;

/**
 * wang 2018/3/8
 */
public class Test {

  private static final String filePath = "F:\\data_doc\\data\\test.xlsx";

  public static void main(String[] args) {

    File file = new File(filePath);
    ArrayList<ArrayList<Object>> result = ExcelUtil.readExcel(file);
    for (int i = 0; i < result.size(); i++) {
      for (int j = 0; j < result.get(i).size(); j++) {
        System.out.println(i + "行 " + j + "列  " + result.get(i).get(j).toString());
      }
    }
    ExcelUtil.writeExcel(result, "F:\\data_doc\\data\\bb.xls");
  }

}
