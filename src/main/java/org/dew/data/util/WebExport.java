package org.dew.data.util;

import java.io.IOException;
import java.io.OutputStream;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import javax.servlet.ServletException;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

public 
class WebExport extends HttpServlet
{
  private static final long serialVersionUID = 2186913005374384760L;
  
  public
  void doPost(HttpServletRequest request, HttpServletResponse response)
    throws ServletException, IOException
  {
    doGet(request, response);
  }
  
  public
  void doGet(HttpServletRequest request, HttpServletResponse response)
    throws ServletException, IOException
  {
    String pathInfo  = request.getPathInfo();
    if(pathInfo != null && pathInfo.length() > 0) {
      if(pathInfo.startsWith("/")) pathInfo = pathInfo.substring(1);
    }
    else {
      response.sendError(404); // Not Found
      return;
    }
    
    String entity = pathInfo.toLowerCase();
    String type   = "csv";
    int iSep = pathInfo.indexOf('.');
    if(iSep > 0) {
      entity = pathInfo.substring(0,iSep).toLowerCase();
      type   = pathInfo.substring(iSep+1).toLowerCase();
    }
    
    List<List<Object>> listResult = null;
    
    // Read data
    if(entity.equals("test")) {
      listResult = getTestData();
    }
    
    // Export
    byte[] content = ExportAs.any(listResult, entity, type);
    
    // Send file
    response.setContentType(ExportAs.getContentType(type));
    response.addHeader("content-disposition", "attachment; filename=\"" + entity + "." + type +"\"");
    response.setContentLength(content.length);
    
    OutputStream out = response.getOutputStream();
    out.write(content, 0, content.length);
  }
  
  private static
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
  
  private static
  List<Object> record(Object... objects)
  {
    List<Object> listResult = new ArrayList<Object>(objects.length);
    
    for(int i = 0; i < objects.length; i++) {
      listResult.add(objects[i]);
    }
    
    return listResult;
  }
}
