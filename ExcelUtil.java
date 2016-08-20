
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;


public class ExcelUtil {

  private static final String cellSplitRegex = "(?<=[a-zA-Z])(?=[0-9])";

  /**
   * ExcelUtil的用法:
   * 第一步先取得template 傳入fileName
   * 
   * @param args
   */
  public static void main(String[] args) {

    // for (int i = 0; i < 1000; i++) {
    // System.out.println(toCellName(0, i));
    // }
    try {
      Workbook workbook = loadTemplate("ExcelUtil.xlsx");
      HashMap<String, String> datas = new HashMap<>();
      // datas.put("A1", "Name");
      // datas.put("B1", "Phone");
      // datas.put("A2", "Alan");
      // datas.put("B2", "0987654321");
      // datas.put("A3", "Blan");
      // datas.put("B3", "0912345678");
      datas.put(toCellName(1, 26), "Name");
      writeSheet(workbook.getSheetAt(0), datas);

      FileOutputStream fileOut = new FileOutputStream("ExcelUtil.xlsx");
      workbook.write(fileOut);
    } catch (Exception e) {
      e.printStackTrace();
    }
    System.out.println("write finished!");
  }

  /**
   * 讀取檔案取得template
   * 
   * @param fileName
   * @return
   * @throws Exception
   */
  public static Workbook loadTemplate(String fileName) throws Exception {

    InputStream inp = new FileInputStream(fileName);
    Workbook workbook = WorkbookFactory.create(inp);
    inp.close();
    return workbook;
  }

  /**
   * 讀取檔案取得template
   * 
   * @param fileName
   * @return
   * @throws Exception
   */
  public static Workbook newWorkbook() throws Exception {
    return new XSSFWorkbook();
  }

  /**
   * 傳入一個hashMap cellName對應value
   * 例如:
   * A1:Name, B1:Phone
   * A2:Alan, B2:0987654321
   * 
   * @param sheet
   * @param map
   * @throws Exception
   */
  public static void writeSheet(Sheet sheet, Map<String, String> map) throws Exception {

    for (String cellName : map.keySet()) {
      writeCell(sheet, cellName, map.get(cellName));
    }
  }

  public static void writeCell(Sheet sheet, String cellName, String value) throws Exception {

    List<String> cellNames = Arrays.asList(cellName.split(cellSplitRegex));
    writeCell(sheet, Integer.parseInt(cellNames.get(1)) - 1, translateColumnName(cellNames.get(0)) - 1, value);
  }

  public static void writeCell(Sheet sheet, int rowNumber, int columnNumber, String value) throws Exception {

    Row row = sheet.getRow(rowNumber);
    if (row == null) {
      row = sheet.createRow(rowNumber);
    }

    Cell cell = row.getCell(columnNumber);
    if (cell == null) {
      cell = row.createCell(columnNumber);
    }
    cell.setCellValue(value);
  }

  /**
   * 如果需要取得cellName 如AA12 可以利用這個 傳入rowNumber, columnNumber
   * 如傳入:rowNumber:11,columnNumer:26
   * 將回傳 AA12
   * 可以當作writeSheet中Map的key
   * 
   * @param rowNumber
   * @param columnNumber
   * @return
   */
  public static String toCellName(int rowNumber, int columnNumber) {
    return numberToCharacterRepresentation(columnNumber) + (rowNumber + 1);
  }

  private static String numberToCharacterRepresentation(int number) {
    char[] ls = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".toCharArray();
    String r = "";
    while (true) {
      r = ls[number % 26] + r;
      if (number < 26) {
        break;
      }
      number /= 26;
      number--;
    }
    return r;
  }

  private static int translateColumnName(String columnName) {
    // remove any whitespace
    columnName = columnName.trim();

    StringBuffer buff = new StringBuffer(columnName);

    // string to lower case, reverse then place in char array
    char chars[] = buff.reverse().toString().toLowerCase().toCharArray();

    int retVal = 0, multiplier = 0;

    for (int i = 0; i < chars.length; i++) {
      // retrieve ascii value of character, subtract 96 so number corresponds to
      // place in alphabet. ascii 'a' = 97
      multiplier = (int) chars[i] - 96;
      // mult the number by 26^(position in array)
      retVal += multiplier * Math.pow(26, i);
    }
    return retVal;
  }
}
