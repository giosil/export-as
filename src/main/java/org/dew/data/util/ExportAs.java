package org.dew.data.util;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public
class ExportAs
{
  public static String YES_VALUE      = "X";
  public static String NO_VALUE       = "";
  
  public static int    BOOL_COL_WIDTH = 4000;
  public static int    NUM_COL_WIDTH  = 5000;
  public static int    DATE_COL_WIDTH = 6000;
  public static int    STR_COL_WIDTH  = 8000;
  
  public static String DATE_FORMAT    = "dd/mm/yyyy";
  
  public static String CSV_SEPARATOR  = ";";
  public static String CSV_DELIMITER  = "";
  
  public static String HTML_TH_STYLE  = "background-color:#eeeeee;";
  public static String HTML_TD_STYLE  = "border-color:#cccccc;border-width:1px;";
  
  public static
  byte[] any(List<List<Object>> listData, String title, String type)
  {
    if(type == null || type.length() == 0) {
      return csv(listData);
    }
    
    String typeLC = type.toLowerCase();
    
    if(typeLC.endsWith("xls") || typeLC.endsWith("excel")) {
      
      return xls(listData, title);
    
    }
    else if(typeLC.endsWith("xlsx") || typeLC.endsWith("sheet")) {
      
      return xlsx(listData, title);
      
    }
    else if(typeLC.endsWith("htm") || typeLC.endsWith("html")) {
      
      return html(listData, title);
      
    }
    else if(typeLC.endsWith("json") || typeLC.endsWith("js")) {
      
      return json(listData, title);
      
    }
    else if(typeLC.endsWith("pdf")) {
      
      return pdf(listData, title);
      
    }
    
    return csv(listData);
  }
  
  public static
  byte[] xls(List<List<Object>> listData, String title)
  {
    Workbook workBook = new HSSFWorkbook();
    
    return fill(workBook, listData, title);
  }
  
  public static
  byte[] xlsx(List<List<Object>> listData, String title)
  {
    Workbook workBook = new XSSFWorkbook();
    
    return fill(workBook, listData, title);
  }
  
  public static
  byte[] csv(List<List<Object>> listData)
  {
    if(listData == null || listData.size() == 0) {
      return "".getBytes();
    }
    
    StringBuilder sb = new StringBuilder(listData.size() * 10);
    for(int i = 0; i < listData.size(); i++) {
      List<Object> listRecord = listData.get(i);
      if(listRecord == null) continue;
      
      sb.append(csvRow(listRecord));
      sb.append((char) 13);
      sb.append((char) 10);
    }
    
    return sb.toString().getBytes();
  }
  
  public static
  byte[] csv(List<List<Object>> listData, String title)
  {
    if(listData == null || listData.size() == 0) {
      return "".getBytes();
    }
    
    StringBuilder sb = new StringBuilder(listData.size() * 10);
    for(int i = 0; i < listData.size(); i++) {
      List<Object> listRecord = listData.get(i);
      if(listRecord == null) continue;
      
      sb.append(csvRow(listRecord));
      sb.append((char) 13);
      sb.append((char) 10);
    }
    
    return sb.toString().getBytes();
  }
  
  public static
  byte[] json(List<List<Object>> listData)
  {
    if(listData == null || listData.size() == 0) {
      return "[]".getBytes();
    }
    
    int row = 0;
    
    StringBuilder sb = new StringBuilder(listData.size() * 10);
    sb.append('[');
    for(int i = 0; i < listData.size(); i++) {
      List<Object> listRecord = listData.get(i);
      if(listRecord == null) continue;
      
      StringBuilder sbRow = new StringBuilder();
      for(int c = 0; c < listRecord.size(); c++) {
        Object value = listRecord.get(c);
        sbRow.append("," + jsonValue(value));
      }
      
      if(row > 0) sb.append(',');
      row++;
      
      String sRow = sbRow.length() > 0 ? sbRow.substring(1) : "";
      sb.append("[" + sRow + "]");
    }
    sb.append(']');
    
    return sb.toString().getBytes();
  }
  
  public static
  byte[] json(List<List<Object>> listData, String title)
  {
    if(title == null || title.length() == 0) {
      return json(listData);
    }
    
    if(listData == null || listData.size() == 0) {
      return ("{" + jsonValue(title) + ":[]}").getBytes();
    }
    
    int row = 0;
    
    StringBuilder sb = new StringBuilder(listData.size() * 10);
    sb.append("{" + jsonValue(title) + ":[");
    for(int i = 0; i < listData.size(); i++) {
      List<Object> listRecord = listData.get(i);
      if(listRecord == null) continue;
      
      StringBuilder sbRow = new StringBuilder();
      for(int c = 0; c < listRecord.size(); c++) {
        Object value = listRecord.get(c);
        sbRow.append("," + jsonValue(value));
      }
      
      if(row > 0) sb.append(',');
      row++;
      
      String sRow = sbRow.length() > 0 ? sbRow.substring(1) : "";
      sb.append("[" + sRow + "]");
    }
    sb.append("]}");
    
    return sb.toString().getBytes();
  }
  
  public static
  byte[] html(List<List<Object>> listData, String title)
  {
    if(listData == null) {
      listData = new ArrayList<List<Object>>(0);
    }
    
    String pageTitle = "";
    if(title != null && title.length() > 1) {
      pageTitle = "<h2>" + title.substring(0,1).toUpperCase() + title.substring(1) + "</h2>";
    }
    
    if(HTML_TH_STYLE == null) HTML_TH_STYLE = "";
    if(HTML_TD_STYLE == null) HTML_TD_STYLE = "";
    
    int columns  = 0;
    String style = "";
    // Header
    StringBuilder sb = new StringBuilder(listData.size() * 30);
    if(listData.size() > 0) {
      List<Object> listHeader = listData.get(0);
      columns = listHeader != null ? listHeader.size() : 0;
      if(columns < 20) style="<style>table,th,td{font-size:11px;" + HTML_TD_STYLE + "}th{" + HTML_TH_STYLE + "}</style>"; else
      if(columns < 30) style="<style>table,th,td{font-size:10px;" + HTML_TD_STYLE + "}th{" + HTML_TH_STYLE + "}</style>"; else
      if(columns < 40) style="<style>table,th,td{font-size:9px;"  + HTML_TD_STYLE + "}th{" + HTML_TH_STYLE + "}</style>";  else
      if(columns < 50) style="<style>table,th,td{font-size:8px;"  + HTML_TD_STYLE + "}th{" + HTML_TH_STYLE + "}</style>";  else {
        style="<style>table,th,td{font-size:7px;" + HTML_TD_STYLE + "}th{" + HTML_TH_STYLE + "}</style>";
      }
      sb.append("<html>" + style + "<body>" + pageTitle + "<table border=\"1\" cellspacing=\"0\" width=\"100%\">");
      sb.append(htmlTableRow(listHeader, "th"));
    }
    // Body
    for(int i = 1; i < listData.size(); i++) {
      List<Object> listRecord = listData.get(i);
      if(listRecord == null) continue;
      
      sb.append(htmlTableRow(listRecord, "td"));
    }
    sb.append("</table></body></html>");
    
    return sb.toString().getBytes();
  }
  
  public static
  byte[] pdf(List<List<Object>> listData, String title)
  {
    byte[] htmlContent = html(listData, title);
    
    ByteArrayOutputStream result = new ByteArrayOutputStream();
    
    com.itextpdf.text.Document pdfDocument = new com.itextpdf.text.Document();
    try {
      pdfDocument.setMargins(4, 4, 4, 4);
      
      com.itextpdf.text.pdf.PdfWriter writer = com.itextpdf.text.pdf.PdfWriter.getInstance(pdfDocument, result);
      pdfDocument.open();
      
      com.itextpdf.tool.xml.XMLWorkerHelper.getInstance().parseXHtml(writer, pdfDocument, new ByteArrayInputStream(htmlContent));
    }
    catch(Exception ex) {
      System.err.println("ExportAs.pdf: " + ex);
    }
    finally {
      if(pdfDocument != null) try { pdfDocument.close(); } catch(Exception ex) {}
    }
    
    return result.toByteArray();
  }
  
  public static
  List<List<Object>> data(byte[] xlsx)
  {
    List<List<Object>> listResult = new ArrayList<List<Object>>();
    
    if(xlsx == null || xlsx.length < 3) {
      return listResult;
    }
    
    Workbook workbook = null;
    try {
      if(xlsx[0] == 'P') {
        workbook = new XSSFWorkbook(new ByteArrayInputStream(xlsx));
      }
      else {
        workbook = new HSSFWorkbook(new ByteArrayInputStream(xlsx));
      }
      
      Sheet sheet0 = workbook.getSheetAt(0);
      
      Iterator<Row> rowIterator = sheet0.iterator();
      
      while(rowIterator.hasNext()) {
        Row row = rowIterator.next();
        
        List<Object> record = new ArrayList<Object>();
        listResult.add(record);
        
        Iterator<Cell> cellIterator = row.iterator();
        while(cellIterator.hasNext()) {
          Cell cell = cellIterator.next();
          
          switch (cell.getCellType()) {
          case BLANK:
            record.add("");
            break;
          case BOOLEAN:
            record.add(cell.getBooleanCellValue());
            break;
          case NUMERIC:
            CellStyle cellStyle = cell.getCellStyle();
            if(cellStyle != null) {
              String dataFormat = cellStyle.getDataFormatString();
              if(dataFormat != null && dataFormat.length() > 0) {
                if(dataFormat.indexOf('/') > 0 || dataFormat.indexOf('-') > 0) {
                  record.add(cell.getDateCellValue());
                  continue;
                }
              }
            }
            record.add(cell.getNumericCellValue());
            break;
          default:
            record.add(cell.getStringCellValue());
            break;
          }
        }
      }
    }
    catch(Exception ex) {
      System.err.println("ExportAs.data: " + ex);
    }
    finally {
      if(workbook != null) try { workbook.close(); } catch(Exception ex) {}
    }
    
    return listResult;
  }
  
  public static
  String getContentType(Object fileName)
  {
    if(fileName == null) {
      return "text/plain";
    }
    String filename = fileName.toString();
    String ext = "";
    int dot = filename.lastIndexOf('.');
    if(dot >= 0 && dot < filename.length() - 1) {
      ext = filename.substring(dot + 1).toLowerCase();
    }
    else if(filename.length() < 5) {
      ext = filename.toLowerCase();
    }
    if(ext.equals("txt"))  return "text/plain";
    if(ext.equals("dat"))  return "text/plain";
    if(ext.equals("csv"))  return "text/plain";
    if(ext.equals("html")) return "text/html";
    if(ext.equals("htm"))  return "text/html";
    if(ext.equals("json")) return "application/json";
    if(ext.equals("xml"))  return "text/xml";
    if(ext.equals("log"))  return "text/plain";
    if(ext.equals("rtf"))  return "application/rtf";
    if(ext.equals("doc"))  return "application/msword";
    if(ext.equals("docx")) return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    if(ext.equals("xlsx")) return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    if(ext.equals("xls"))  return "application/x-msexcel";
    if(ext.equals("pdf"))  return "application/pdf";
    if(ext.equals("gif"))  return "image/gif";
    if(ext.equals("bmp"))  return "image/bmp";
    if(ext.equals("jpg"))  return "image/jpeg";
    if(ext.equals("jpeg")) return "image/jpeg";
    if(ext.equals("tif"))  return "image/tiff";
    if(ext.equals("tiff")) return "image/tiff";
    if(ext.equals("png"))  return "image/png";
    if(ext.equals("mpg"))  return "video/mpeg";
    if(ext.equals("mpeg")) return "video/mpeg";
    if(ext.equals("mp4"))  return "video/mpeg";
    if(ext.equals("mp3"))  return "audio/mp3";
    if(ext.equals("wav"))  return "audio/wav";
    if(ext.equals("wma"))  return "audio/wma";
    if(ext.equals("mov"))  return "video/quicktime";
    if(ext.equals("tar"))  return "application/x-tar";
    if(ext.equals("zip"))  return "application/x-zip-compressed";
    return "application/" + ext;
  }
  
  private static
  byte[] fill(Workbook workBook, List<List<Object>> listData, String title)
  {
    ByteArrayOutputStream result = new ByteArrayOutputStream();
    
    if(title == null || title.length() == 0) {
      title = "export";
    }
    
    Map<String, CellStyle> mapStyles = createStyles(workBook);
    
    Sheet sheet = workBook.createSheet(title);
    
    if(listData == null || listData.size() == 0) {
      try {
        workBook.write(result);
      }
      catch(Exception ex) {
        System.err.println("ExportAs.excel: " + ex);
      }
      return result.toByteArray();
    }
    
    Row row = sheet.createRow(0);
    
    // Header
    List<Object> listHeader = listData.get(0);
    for(int c = 0; c < listHeader.size(); c++) {
      createCell(sheet, row, c, mapStyles.get("headerl"), listHeader.get(c));
    }
    // Set Column Width
    if(listData.size() > 1) {
      List<Object> listFirstRow = listData.get(1);
      if(listFirstRow == null) listFirstRow = new ArrayList<Object>(0);
      for(int c = 0; c < listHeader.size(); c++) {
        if(c < listFirstRow.size()) {
          Object value = listFirstRow.get(c);
          if(value instanceof Number) {
            sheet.setColumnWidth(c, NUM_COL_WIDTH);
          }
          else if(value instanceof Boolean) {
            sheet.setColumnWidth(c, BOOL_COL_WIDTH);
          }
          else if(value instanceof Date) {
            sheet.setColumnWidth(c, DATE_COL_WIDTH);
          }
          else if(value instanceof Calendar) {
            sheet.setColumnWidth(c, DATE_COL_WIDTH);
          }
          else {
            sheet.setColumnWidth(c, STR_COL_WIDTH);
          }
        }
        else {
          sheet.setColumnWidth(c, NUM_COL_WIDTH);
        }
      }
    }
    else {
      for(int c = 0; c < listHeader.size(); c++) {
        sheet.setColumnWidth(c, NUM_COL_WIDTH);
      }
    }
    
    // Body
    for(int r = 1; r < listData.size(); r++) {
      
      row = sheet.createRow(r);
      
      List<Object> listRecord = listData.get(r);
      if(listRecord == null) continue;
      
      for(int c = 0; c < listRecord.size(); c++) {
        Object value = listRecord.get(c);
        if(value instanceof Number) {
          // Right Horizontal Alignment (r)
          createCell(sheet, row, c, mapStyles.get("whiter"), value);
        }
        else if(value instanceof Boolean) {
          // Center Horizontal Alignment (c)
          createCell(sheet, row, c, mapStyles.get("whitec"), value);
        }
        else if(value instanceof Date) {
          // Left Horizontal Alignment width Date Format (d)
          createCell(sheet, row, c, mapStyles.get("whited"), value);
        }
        else if(value instanceof Calendar) {
          // Left Horizontal Alignment width Date Format (d)
          createCell(sheet, row, c, mapStyles.get("whited"), value);
        }
        else {
          // Left Horizontal Alignment (l)
          createCell(sheet, row, c, mapStyles.get("whitel"), value);
        }
      }
    }
    
    try {
      workBook.write(result);
    }
    catch(Exception ex) {
      System.err.println("ExportAs.excel: " + ex);
    }
    return result.toByteArray();
  }
  
  private static
  Cell createCell(Sheet sheet, Row row, int iCol, CellStyle cellStyle, Object value)
  {
    Cell cell = row.createCell(iCol);
    cell.setCellStyle(cellStyle);
    
    if(value == null) {
      cell.setCellValue("");
    }
    else if(value instanceof Date) {
      Calendar cal = Calendar.getInstance();
      cal.setTimeInMillis(((Date) value).getTime());
      cell.setCellValue(cal);
    }
    else if(value instanceof Calendar) {
      cell.setCellValue((java.util.Calendar) value);
    }
    else if(value instanceof Number) {
      cell.setCellValue(((Number) value).doubleValue());
    }
    else if(value instanceof Boolean) {
      cell.setCellValue(((Boolean) value).booleanValue() ? YES_VALUE : NO_VALUE);
    }
    else {
      cell.setCellValue(value.toString().trim());
    }
    
    return cell;
  }
  
  private static 
  Map<String, CellStyle> createStyles(Workbook workBook)
  {
    Map<String, CellStyle> mapResult = new HashMap<String, CellStyle>();
    
    mapResult.put("whiteb",  createStyle(workBook, true,  HorizontalAlignment.CENTER,   false));
    mapResult.put("whitec",  createStyle(workBook, false, HorizontalAlignment.CENTER,   false));
    mapResult.put("whitel",  createStyle(workBook, false, HorizontalAlignment.LEFT,     false));
    mapResult.put("whiter",  createStyle(workBook, false, HorizontalAlignment.RIGHT,    false));
    mapResult.put("whited",  createStyle(workBook, false, HorizontalAlignment.LEFT,     DATE_FORMAT));
    mapResult.put("headerc", createStyle(workBook, true,  (short) HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex(), HorizontalAlignment.CENTER, true));
    mapResult.put("headerl", createStyle(workBook, true,  (short) HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex(), HorizontalAlignment.LEFT,   true));
    mapResult.put("headerr", createStyle(workBook, true,  (short) HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex(), HorizontalAlignment.RIGHT));
    
    return mapResult;
  }
  
  private static 
  CellStyle createStyle(Workbook workBook, boolean bold, HorizontalAlignment alignment, boolean boWrapText)
  {
    Font font = workBook.createFont();
    font.setFontHeightInPoints((short) 10);
    font.setFontName("Arial");
    font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
    if(bold) font.setBold(bold);
    
    CellStyle cellStyle = workBook.createCellStyle();
    cellStyle.setFont(font);
    cellStyle.setWrapText(boWrapText);
    cellStyle.setAlignment(alignment);
    cellStyle.setBorderBottom(BorderStyle.THIN);
    cellStyle.setBorderTop(BorderStyle.THIN);
    cellStyle.setBorderLeft(BorderStyle.THIN);
    cellStyle.setBorderRight(BorderStyle.THIN);
    return cellStyle;
  }
  
  private static 
  CellStyle createStyle(Workbook workBook, boolean bold, HorizontalAlignment alignment, String dataFormat)
  {
    Font font = workBook.createFont();
    font.setFontHeightInPoints((short) 10);
    font.setFontName("Arial");
    font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
    if(bold) font.setBold(bold);
    
    CellStyle cellStyle = workBook.createCellStyle();
    cellStyle.setFont(font);
    if(dataFormat != null && dataFormat.length() > 0) {
      CreationHelper createHelper = workBook.getCreationHelper();
      cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(dataFormat));
    }
    cellStyle.setAlignment(alignment);
    cellStyle.setBorderBottom(BorderStyle.THIN);
    cellStyle.setBorderTop(BorderStyle.THIN);
    cellStyle.setBorderLeft(BorderStyle.THIN);
    cellStyle.setBorderRight(BorderStyle.THIN);
    return cellStyle;
  }
  
  private static 
  CellStyle createStyle(Workbook workBook, boolean bold, short background, HorizontalAlignment alignment)
  {
    Font font = workBook.createFont();
    font.setFontHeightInPoints((short) 10);
    font.setFontName("Arial");
    font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
    if(bold) font.setBold(bold);
    
    CellStyle cellStyle = workBook.createCellStyle();
    cellStyle.setFont(font);
    cellStyle.setWrapText(false);
    cellStyle.setFillForegroundColor((short) background);
    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    cellStyle.setAlignment(alignment);
    cellStyle.setBorderBottom(BorderStyle.THIN);
    cellStyle.setBorderTop(BorderStyle.THIN);
    cellStyle.setBorderLeft(BorderStyle.THIN);
    cellStyle.setBorderRight(BorderStyle.THIN);
    return cellStyle;
  }
  
  private static 
  CellStyle createStyle(Workbook workBook, boolean bold, short background, HorizontalAlignment alignment, boolean boWrapText)
  {
    Font font = workBook.createFont();
    font.setFontHeightInPoints((short) 10);
    font.setFontName("Arial");
    font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
    if(bold) font.setBold(bold);
    
    CellStyle cellStyle = workBook.createCellStyle();
    cellStyle.setFont(font);
    cellStyle.setWrapText(boWrapText);
    cellStyle.setFillForegroundColor((short) background);
    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    cellStyle.setAlignment(alignment);
    cellStyle.setBorderBottom(BorderStyle.THIN);
    cellStyle.setBorderTop(BorderStyle.THIN);
    cellStyle.setBorderLeft(BorderStyle.THIN);
    cellStyle.setBorderRight(BorderStyle.THIN);
    return cellStyle;
  }
  
  private static
  String csvRow(List<?> items)
  {
    if(items == null || items.size() == 0) {
      return "";
    }
    
    if(CSV_SEPARATOR == null) CSV_SEPARATOR = "";
    if(CSV_DELIMITER == null) CSV_DELIMITER = "";
    
    String result = "";
    for(Object item : items) {
      if(item instanceof Date) {
        result += CSV_SEPARATOR + CSV_DELIMITER + formatDate((Date) item) + CSV_DELIMITER;
      }
      else if(item instanceof Calendar) {
        result += CSV_SEPARATOR + CSV_DELIMITER + formatDate((Calendar) item) + CSV_DELIMITER;
      }
      else if(item instanceof Boolean) {
        String yesNo = ((Boolean) item).booleanValue() ? YES_VALUE : NO_VALUE;
        result += CSV_SEPARATOR + CSV_DELIMITER + yesNo + CSV_DELIMITER;
      }
      else if(item instanceof Number) {
        result += CSV_SEPARATOR + CSV_DELIMITER + item.toString().replace('.', ',') + CSV_DELIMITER;
      }
      else if(item != null) {
        result += CSV_SEPARATOR + CSV_DELIMITER + item.toString().replace(';', ',').replace('\n', ' ').replace("\r", "").replace('"', '\'').trim() + CSV_DELIMITER;
      }
      else {
        result += CSV_SEPARATOR + CSV_DELIMITER + CSV_DELIMITER;
      }
    }
    if(result.length() > 0) {
      result = result.substring(CSV_SEPARATOR.length());
    }
    return result;
  }
  
  private static
  String htmlTableRow(List<?> items, String tag)
  {
    if(tag == null) tag = "td";
    if(tag.startsWith("<") && tag.endsWith(">")) tag = tag.substring(1, tag.length()-1);
    if(tag.length() < 2) tag = "td";
    if(items == null || items.size() == 0) {
      return "<tr><" + tag + "></" + tag + "></tr>";
    }
    StringBuilder sb = new StringBuilder();
    sb.append("<tr>");
    for(Object item : items) {
      String align = "";
      if(item instanceof Number) {
        align = " align=\"right\"";
      }
      else
      if(item instanceof Boolean) {
        align = " align=\"center\"";
      }
      sb.append("<" + tag + align + ">");
      if(item instanceof Date) {
        sb.append(formatDate((Date) item));
      }
      else if(item instanceof Calendar) {
        sb.append(formatDate((Calendar) item));
      }
      else if(item instanceof Boolean) {
        sb.append(((Boolean) item).booleanValue() ? YES_VALUE : NO_VALUE);
      }
      else if(item != null) {
        sb.append(item.toString().replace("<", "&lt;").replace(">", "&gt;"));
      }
      else {
        sb.append("&nbsp;");
      }
      sb.append("</" + tag + ">");
    }
    sb.append("</tr>");
    return sb.toString();
  }
  
  private static 
  String jsonValue(Object value) 
  {
    if(value == null) {
      return "null";
    }
    if(value instanceof Number) {
      return value.toString();
    }
    if(value instanceof Boolean) {
      return value.toString();
    }
    if(value instanceof Date) {
      return "\"" + formatDate((Date) value) + "\"";
    }
    if(value instanceof Calendar) {
      return "\"" + formatDate((Calendar) value) + "\"";
    }
    String string = value.toString();
    if(string.length() == 0) {
      return "\"\"";
    }
    char b;
    char c = 0;
    String hhhh;
    int i;
    int len = string.length();
    StringBuilder sb = new StringBuilder(len + 2);
    sb.append('"');
    for(i = 0; i < len; i += 1) {
      b = c;
      c = string.charAt(i);
      switch(c) {
        case '\\':
        case '"':
          sb.append('\\');
          sb.append(c);
        break;
        case '/':
          if(b == '<') {
            sb.append('\\');
          }
          sb.append(c);
        break;
        case '\b':
          sb.append("\\b");
        break;
        case '\t':
          sb.append("\\t");
        break;
        case '\n':
          sb.append("\\n");
        break;
        case '\f':
          sb.append("\\f");
        break;
        case '\r':
          sb.append("\\r");
        break;
        default:
          if(c < ' ' ||(c >= '\u0080' && c < '\u00a0') ||(c >= '\u2000' && c < '\u2100')) {
            sb.append("\\u");
            hhhh = Integer.toHexString(c);
            sb.append("0000", 0, 4 - hhhh.length());
            sb.append(hhhh);
          } 
          else {
            sb.append(c);
          }
      }
    }
    sb.append('"');
    return sb.toString();
  }
  
  private static
  String formatDate(Date date)
  {
    if(date == null) return "";
    
    Calendar cal = Calendar.getInstance();
    cal.setTimeInMillis(date.getTime());
    
    return formatDate(cal);
  }
  
  private static
  String formatDate(Calendar cal)
  {
    if(cal == null) return "";
    
    int iYear  = cal.get(Calendar.YEAR);
    int iMonth = cal.get(Calendar.MONTH) + 1;
    int iDay   = cal.get(Calendar.DATE);
    
    String sYear  = String.valueOf(iYear);
    String sMonth = iMonth < 10 ? "0" + iMonth : String.valueOf(iMonth);
    String sDay   = iDay   < 10 ? "0" + iDay   : String.valueOf(iDay);
    if(iYear < 10) {
      sYear = "000" + sYear;
    }
    else if(iYear < 100) {
      sYear = "00" + sYear;
    }
    else if(iYear < 1000) {
      sYear = "0" + sYear;
    }
    
    if(DATE_FORMAT == null || DATE_FORMAT.length() == 0 || DATE_FORMAT.startsWith("#")) {
      return sYear + sMonth + sDay;
    }
    else if(DATE_FORMAT.indexOf('-') >= 0 || DATE_FORMAT.startsWith("y")) {
      return sYear + "-" + sMonth + "-" + sDay;
    }
    else if(DATE_FORMAT.startsWith("m") || DATE_FORMAT.startsWith("u")) {
      return sMonth + "/" + sDay + "/" + sYear;
    }
    return sDay + "/" + sMonth + "/" + sYear;
  }
}
