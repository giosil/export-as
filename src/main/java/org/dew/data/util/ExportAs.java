package org.dew.data.util;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.tool.xml.XMLWorkerHelper;

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
  public static String YES_VALUE     = "X";
  public static String NO_VALUE      = " ";
  public static int    COL_WIDTH     = 8000;
  public static String DATE_FORMAT   = "dd/mm/yyyy";
  public static char   CSV_SEPARATOR = ';';
  
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
    else if(typeLC.endsWith("pdf")) {
      
      return pdf(listData, title);
      
    }
    
    return csv(listData);
  }
  
  public static
  byte[] xls(List<List<Object>> listData, String title)
  {
    ByteArrayOutputStream result = new ByteArrayOutputStream();
    
    if(title == null || title.length() == 0) {
      title = "export";
    }
    
    Workbook workBook = new HSSFWorkbook();
    
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
    for(int c = 0; c < listHeader.size(); c++) {
      sheet.setColumnWidth(c, COL_WIDTH);
    }
    
    // Body
    for(int r = 1; r < listData.size(); r++) {
      
      row = sheet.createRow(r);
      
      List<Object> listRecord = listData.get(r);
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
  
  public static
  byte[] xlsx(List<List<Object>> listData, String title)
  {
    ByteArrayOutputStream result = new ByteArrayOutputStream();
    
    if(title == null || title.length() == 0) {
      title = "export";
    }
    
    Workbook workBook = new XSSFWorkbook();
    
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
    for(int c = 0; c < listHeader.size(); c++) {
      sheet.setColumnWidth(c, COL_WIDTH);
    }
    
    // Body
    for(int r = 1; r < listData.size(); r++) {
      
      row = sheet.createRow(r);
      
      List<Object> listRecord = listData.get(r);
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
  
  public static
  byte[] csv(List<List<Object>> listData)
  {
    if(listData == null || listData.size() == 0) {
      return "".getBytes();
    }
    
    StringBuilder sb = new StringBuilder(listData.size() * 10);
    for(int i = 0; i < listData.size(); i++) {
      List<Object> listRecord = listData.get(i);
      
      sb.append(toCSV(listRecord));
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
      
      sb.append(toCSV(listRecord));
      sb.append((char) 13);
      sb.append((char) 10);
    }
    
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
    
    int columns  = 0;
    String style = "";
    // Header
    StringBuilder sb = new StringBuilder();
    if(listData.size() > 0) {
      List<Object> listHeader = listData.get(0);
      columns = listHeader != null ? listHeader.size() : 0;
      if(columns < 20) style="<style>table,th,td{font-size:11px;border-color:#cccccc;border-width:1px;}th{background-color:#eeeeee;}</style>"; else
      if(columns < 30) style="<style>table,th,td{font-size:10px;border-color:#cccccc;border-width:1px;}th{background-color:#eeeeee;}</style>"; else
      if(columns < 40) style="<style>table,th,td{font-size:9px;border-color:#cccccc;border-width:1px;}th{background-color:#eeeeee;}</style>";  else
      if(columns < 50) style="<style>table,th,td{font-size:8px;border-color:#cccccc;border-width:1px;}th{background-color:#eeeeee;}</style>";  else {
        style="<style>table,th,td{font-size:7px;border-color:#cccccc;border-width:1px;}th{background-color:#eeeeee;}</style>";
      }
      sb.append("<html>" + style + "<body>" + pageTitle + "<table border=\"1\" cellspacing=\"0\" width=\"100%\">");
      sb.append(htmlTableRow(listHeader, "th"));
    }
    // Body
    for(int i = 1; i < listData.size(); i++) {
      List<Object> listRecord = listData.get(i);
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
      
      XMLWorkerHelper.getInstance().parseXHtml(writer, pdfDocument, new ByteArrayInputStream(htmlContent));
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
  String getContentType(Object fileName)
  {
    if(fileName == null) return "text/plain";
    String filename = fileName.toString();
    String ext = "";
    int dot = filename.lastIndexOf('.');
    if(dot >= 0 && dot < filename.length() - 1) {
      ext = filename.substring(dot + 1).toLowerCase();
    }
    else if(filename.length() < 5) {
      ext = filename.toLowerCase();
    }
    if(ext.equals("txt"))   return "text/plain";  else
    if(ext.equals("dat"))   return "text/plain";  else
    if(ext.equals("csv"))   return "text/plain";  else
    if(ext.equals("html"))  return "text/html";   else
    if(ext.equals("htm"))   return "text/html";   else
    if(ext.equals("xml"))   return "text/xml";    else
    if(ext.equals("log"))   return "text/plain";  else
    if(ext.equals("rtf"))   return "application/rtf"; else
    if(ext.equals("doc"))   return "application/msword"; else
    if(ext.equals("docx"))  return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"; else
    if(ext.equals("xls"))   return "application/x-msexcel"; else
    if(ext.equals("xlsx"))  return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; else
    if(ext.equals("pdf"))   return "application/pdf"; else
    if(ext.equals("gif"))   return "image/gif";   else
    if(ext.equals("bmp"))   return "image/bmp";   else
    if(ext.equals("jpg"))   return "image/jpeg";  else
    if(ext.equals("jpeg"))  return "image/jpeg";  else
    if(ext.equals("tif"))   return "image/tiff";  else
    if(ext.equals("tiff"))  return "image/tiff";  else
    if(ext.equals("png"))   return "image/png";   else
    if(ext.equals("mpg"))   return "video/mpeg";  else
    if(ext.equals("mpeg"))  return "video/mpeg";  else
    if(ext.equals("mp4"))   return "video/mpeg";  else
    if(ext.equals("mp3"))   return "audio/mp3";   else
    if(ext.equals("wav"))   return "audio/wav";   else
    if(ext.equals("wma"))   return "audio/wma";   else
    if(ext.equals("mov"))   return "video/quicktime";   else
    if(ext.equals("tar"))   return "application/x-tar"; else
    if(ext.equals("zip"))   return "application/x-zip-compressed"; else
    return "application/" + ext;
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
  String toCSV(List<?> items)
  {
    if(items == null || items.size() == 0) {
      return "";
    }
    
    String result = "";
    for(Object item : items) {
      if(item instanceof Date) {
        Calendar cal = Calendar.getInstance();
        cal.setTimeInMillis(((Date) item).getTime());
        result += CSV_SEPARATOR + formatDate(cal);
      }
      else if(item instanceof Calendar) {
        result += CSV_SEPARATOR + formatDate((Calendar) item);
      }
      else if(item instanceof Boolean) {
        String yesNo = ((Boolean) item).booleanValue() ? YES_VALUE : NO_VALUE.trim();
        result += CSV_SEPARATOR + yesNo;
      }
      else if(item instanceof Number) {
        result += CSV_SEPARATOR + item.toString().replace('.', ',');
      }
      else if(item != null) {
        result += CSV_SEPARATOR + item.toString().replace(';', ',').replace('\n', ' ').replace("\r", "").replace('"', '\'').trim();
      }
      else {
        result += CSV_SEPARATOR;
      }
    }
    if(result.length() > 0) result = result.substring(1);
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
        Calendar cal = Calendar.getInstance();
        cal.setTimeInMillis(((Date) item).getTime());
        sb.append(formatDate(cal));
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
    
    if(DATE_FORMAT == null || DATE_FORMAT.length() == 0) {
      return sYear + sMonth + sDay;
    }
    else if(DATE_FORMAT.indexOf('-') >= 0) {
      return sYear + "-" + sMonth + "-" + sDay;
    }
    else if(DATE_FORMAT.startsWith("m")) {
      return sMonth + "/" + sDay + "/" + sYear;
    }
    return sDay + "/" + sMonth + "/" + sYear;
  }
}
