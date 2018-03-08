package com.github.when.utis.web;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * https://github.com/joewoo999/selenium_webuitest/blob/master/src/main/java/com/github/joseph/utils/MSExcelUtils.java
 * 用于处理Excel读写操作 wang 2018/3/8
 */
public class MSExcelUtils {

  private Workbook workbook;
  private File excelFile;
  public static int EXCEL97 = 0;
  public static int EXCEL2007 = 1;

  public MSExcelUtils() {
  }

  public MSExcelUtils(String filePath) {
    this.open(filePath);
  }

  public File getExcelFile() {
    return excelFile;
  }

  public Workbook getWorkbook() {
    return workbook;
  }

  /**
   * 打开工作簿
   */
  public void open(String path) {
    this.open(new File(path));
  }

  /**
   * 打开工作簿
   */
  public void open(File file) {
    this.excelFile = file;
    FileInputStream fis = null;
    try {
      fis = new FileInputStream(file);
      workbook = WorkbookFactory.create(fis);
      this.init();
    } catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
      throw new RuntimeException(e);
    } finally {
      this.closeStream(fis, true);
    }
  }

  /**
   * 初始化工作簿
   */
  private void init() {
    Sheet sheet = workbook.getSheetAt(0);
    Row row = sheet.getRow(0);
    if (row == null) {
      row = sheet.createRow(0);
    }
    Cell cell = row.getCell(0);
    if (cell == null) {
      cell = row.createCell(0);
    }
    CellStyle defaultStyle = cell.getCellStyle();
    String cellValue = cell.getStringCellValue();
    cell.setCellStyle(defaultStyle);
    cell.setCellValue(cellValue);
    this.reEvaluateBook();
    // this.save();

  }

  /**
   * 重新计算工作簿内公式
   */
  public void reEvaluateBook() {
    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
    evaluator.clearAllCachedResultValues(); // 清空原结果
    evaluator.evaluateAll();
  }

  /**
   * 当前工作簿下打开工作表
   *
   * @param sheetName 工作表名
   */
  public MSExcelSheet sheet(String sheetName) {
    if (workbook == null) {
      throw new NullPointerException("还未打开工作簿,无法读取工作表.");
    }
    Sheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      throw new NullPointerException("工作表: [" + sheetName + "]不存在.");
    }
    return new MSExcelSheet(sheet);
  }

  /**
   * 当前操作工作簿下打开工作表
   *
   * @param index 工作表位置索引
   */
  public MSExcelSheet sheet(int index) {
    if (workbook == null) {
      throw new NullPointerException("还未打开工作簿,无法读取工作表.");
    }
    Sheet sheet = workbook.getSheetAt(index);
    if (sheet == null) {
      throw new NullPointerException("工作表: [" + index + "]不存在.");
    }
    return new MSExcelSheet(sheet);
  }

  /**
   * 将当前工作簿写入文件
   */
  public void save() {
    this.saveAs(excelFile);
  }

  /**
   * 将当前工作簿写入文件
   */
  public void saveAs(String filePath) {
    this.saveAs(new File(filePath));
  }

  /**
   * 将当前工作簿写入文件
   */
  public void saveAs(File file) {
    FileOutputStream fos = null;
    try {
      fos = new FileOutputStream(file);
      workbook.write(fos);
    } catch (IOException e) {
      throw new RuntimeException(e);
    } finally {
      this.closeStream(fos, true);
    }
  }

  /**
   * excel版本
   *
   * @return 0->EXCEL97, 1->EXCEL2007
   */
  public int getVersion() {
    FileInputStream fis = null;
    try {
      fis = new FileInputStream(excelFile);
      if (NPOIFSFileSystem.hasPOIFSHeader(fis)) {
        return EXCEL97;
      } else if (POIXMLDocument.hasOOXMLHeader(fis)) {
        return EXCEL2007;
      } else {
        throw new RuntimeException("打开的excel文件非MS_97或MS_2007版本.");
      }
    } catch (IOException e) {
      throw new RuntimeException(e);
    } finally {
      this.closeStream(fis, true);
    }
  }

  /**
   * close InputStream
   */
  private void closeStream(FileInputStream stream, boolean success) {
    try {
      if (null != stream) {
        stream.close();
      }
    } catch (IOException e) {
      if (success) {
        throw new RuntimeException(e);
      }
      e.printStackTrace();
    }
  }

  /**
   * close OutputStream
   */
  private void closeStream(FileOutputStream stream, boolean success) {
    try {
      if (null != stream) {
        stream.close();
      }
    } catch (IOException e) {
      if (success) {
        throw new RuntimeException(e);
      }
      e.printStackTrace();
    }
  }

  public class MSExcelSheet {

    private Sheet sheet;

    public MSExcelSheet(Sheet sheet) {
      this.sheet = sheet;
    }

    public Sheet getSheet() {
      return sheet;
    }

    /**
     * 获取工作表行
     *
     * @param index 行
     * @return 行
     */
    public Row row(int index) {
      return row(index, false);
    }

    /**
     * 获取工作表行
     *
     * @param index 行
     * @param isNeedCreated 行不存在时是否创建新的一行
     * @return 行
     */
    public Row row(int index, boolean isNeedCreated) {
      Row row = sheet.getRow(index);
      if (null == row && isNeedCreated) {
        row = sheet.createRow(index);
      }
      return row;
    }

    /**
     * 获取工作表单元格
     *
     * @param rowIndex 行
     * @param colIndex 列
     * @return 单元格
     */
    public MSCell cell(int rowIndex, int colIndex) {
      return cell(rowIndex, colIndex, false);
    }

    /**
     * 获取工作表单元格
     *
     * @param rowIndex 行
     * @param colIndex 列
     * @param isNeedCreated 单元格不存在时是否创建新的单元格
     * @return 单元格
     */
    public MSCell cell(int rowIndex, int colIndex, boolean isNeedCreated) {
      Row row = row(rowIndex, isNeedCreated);
      Cell cell = null;
      if (null != row) {
        cell = row.getCell(colIndex);
        if (null == cell && isNeedCreated) {
          cell = row.createCell(colIndex);
        }
      }
      return new MSCell(cell);
    }
  }

  public class MSCell {

    private Cell cell;

    public MSCell(Cell cell) {
      this.cell = cell;
    }

    public Cell getCell() {
      return cell;
    }

    public boolean isEmpty() {
      return null == cell || Cell.CELL_TYPE_BLANK == cell.getCellType();
    }

    /**
     * 读取单元格内容
     *
     * @return 单元格内容
     */
    public String value() {
      if (isEmpty()) {
        return null;
      }
      int cellType = cell.getCellType();
      switch (cellType) {
        case Cell.CELL_TYPE_FORMULA:
          return getNonFormulaValue(cell, cell.getCachedFormulaResultType()).trim();
        case Cell.CELL_TYPE_NUMERIC:
        case Cell.CELL_TYPE_BOOLEAN:
        case Cell.CELL_TYPE_STRING:
        case Cell.CELL_TYPE_ERROR:
          return getNonFormulaValue(cell, cellType).trim();
        default:
          throw new RuntimeException(String.format("未知的单元格格式: %d", cellType));
      }
    }

    /**
     * 设置单元格值
     *
     * @param value 值
     */
    public void setValue(String value) {
      cell.setCellValue(value);
    }

    /**
     * 设置单元格样式
     *
     * @param cell 单元格
     * @param style 单元格样式
     */
    public void setStyle(Cell cell, CellStyle style) {
      if (null == cell) {
        throw new NullPointerException("不能修改单元格的样式,cell不能为null");
      }
      cell.setCellStyle(style);
    }

    /**
     * 是否为合并单元格
     *
     * @return true or false
     */
    public boolean isMerged() {
      int rIndex = cell.getRowIndex();
      int cIndex = cell.getColumnIndex();
      Sheet sheet = cell.getSheet();
      int mergeCount = sheet.getNumMergedRegions();
      for (int i = 0; i < mergeCount; i++) {
        CellRangeAddress cellRegion = sheet.getMergedRegion(i);
        if (cellRegion.isInRange(rIndex, cIndex)) {
          return true;
        }
      }
      return false;
    }

    /**
     * 获取文本、数字、日期、boolean、error格式单元格内容
     *
     * @param cell 单元格
     * @param cellType 单元格格式类型
     * @return 单元格内容
     */
    private String getNonFormulaValue(Cell cell, int cellType) {
      switch (cellType) {
        case Cell.CELL_TYPE_NUMERIC:
          return getDateOrNumValue(cell);
        case Cell.CELL_TYPE_BOOLEAN:
          return String.valueOf(cell.getBooleanCellValue());
        case Cell.CELL_TYPE_STRING:
          return cell.getStringCellValue();
        case Cell.CELL_TYPE_ERROR:
          return ErrorEval.getText(cell.getErrorCellValue());
        default:
          throw new RuntimeException(String.format("未知的单元格格式类型:%d", cellType));
      }
    }

    /**
     * 获取日期或数字格式的单元格内容
     *
     * @param cell 单元格
     * @return 单元格内容
     */
    private String getDateOrNumValue(Cell cell) {
      if (DateUtil.isCellDateFormatted(cell)) {
        Date date = cell.getDateCellValue();
        return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date);
      } else {
        double value = cell.getNumericCellValue();
        return (value == (long) value) ? String.valueOf((long) value)
            : String.valueOf(value);
      }
    }
  }

}