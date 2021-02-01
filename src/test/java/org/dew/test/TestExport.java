package org.dew.test;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import java.net.URL;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.dew.data.util.ExportAs;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

public class TestExport extends TestCase {
  
  public TestExport(String testName) {
    super(testName);
  }
  
  public static Test suite() {
    return new TestSuite(TestExport.class);
  }
  
  public void testApp() throws Exception {
    
    List<List<Object>> data = getTestData();
    
    byte[] xlsx = ExportAs.xlsx(data, "test");
    
    List<List<Object>> records = ExportAs.data(xlsx);
    
    System.out.println(records);
  }
  
  public static
  List<List<Object>> getTestData()
  {
    List<List<Object>> listResult = new ArrayList<List<Object>>();
    
    // Header
    listResult.add(record("String", "Date",     "Calendar",             "Boolean",    "Integer", "Double"));
    // Body
    listResult.add(record("Text1",  new Date(), Calendar.getInstance(), Boolean.TRUE,  1,        3.14d));
    listResult.add(record("Text2",  new Date(), Calendar.getInstance(), Boolean.FALSE, 2,        6.28d));
    
    return listResult;
  }
  
  public static
  List<Object> record(Object... objects)
  {
    List<Object> listResult = new ArrayList<Object>(objects.length);
    
    for(int i = 0; i < objects.length; i++) {
      listResult.add(objects[i]);
    }
    
    return listResult;
  }
  
  public static
  String getDesktop()
  {
    String sUserHome = System.getProperty("user.home");
    return sUserHome + File.separator + "Desktop";
  }
  
  public static
  String getDesktopPath(String sFileName)
  {
    String sUserHome = System.getProperty("user.home");
    return sUserHome + File.separator + "Desktop" + File.separator + sFileName;
  }
  
  public static
  byte[] readFile(String fileName)
    throws Exception
  {
    if(fileName == null || fileName.length() == 0) {
      return new byte[0];
    }
    
    int iFileSep = fileName.indexOf('/');
    if(iFileSep < 0) iFileSep = fileName.indexOf('\\');
    InputStream is = null;
    if(iFileSep < 0) {
      URL url = Thread.currentThread().getContextClassLoader().getResource(fileName);
      is = url.openStream();
    }
    else {
      is = new FileInputStream(fileName);
    }
    try {
      int n;
      ByteArrayOutputStream baos = new ByteArrayOutputStream();
      byte[] buff = new byte[1024];
      while((n = is.read(buff)) > 0) baos.write(buff, 0, n);
      return baos.toByteArray();
    }
    finally {
      if(is != null) try{ is.close(); } catch(Exception ex) {}
    }
  }
  
  public static
  void saveContent(byte[] content, String sFilePath)
    throws Exception
  {
    if(content == null) return;
    if(content == null || content.length == 0) return;
    File file = new File(sFilePath);
    FileOutputStream fos = null;
    try {
      fos = new FileOutputStream(sFilePath);
      fos.write(content);
      System.out.println(file.getAbsolutePath() + " saved.");
    }
    finally {
      if(fos != null) try{ fos.close(); } catch(Exception ex) {}
    }
  }
}
