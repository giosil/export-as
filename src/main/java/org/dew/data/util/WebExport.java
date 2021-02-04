package org.dew.data.util;

import java.io.IOException;
import java.io.OutputStream;

import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletException;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

public 
class WebExport extends HttpServlet
{
  private static final long serialVersionUID = 8104702489538521122L;

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
    String entity = getEntity(request);
    
    if(entity == null || entity.length() == 0) {
      sendBadRequest(request, response);
      return;
    }
    
    String entityName = getEntityName(entity);
    String entityType = getEntityType(entity);
    
    // Read data
    List<List<Object>> listData = getData(request, entityName, entityType);
    
    if(listData == null) {
      sendEntityNotFound(request, response);
    }
    
    // Export
    byte[] content = ExportAs.any(listData, entityName, entityType);
    
    // Send file
    sendFile(request, response, entityName, entityType, content);
  }
  
  protected
  String getEntity(HttpServletRequest request)
    throws ServletException, IOException
  {
    String pathInfo  = request.getPathInfo();
    
    if(pathInfo == null || pathInfo.length() == 0) {
      return "";
    }
    
    if(pathInfo.startsWith("/")) pathInfo = pathInfo.substring(1);
    
    return pathInfo;
  }
  
  protected
  String getEntityName(String entity)
  {
    if(entity == null || entity.length() == 0) {
      return "";
    }
    int sep = entity.lastIndexOf('.');
    if(sep > 0) {
      return entity.substring(0, sep).toLowerCase();
    }
    return entity.toLowerCase();
  }
  
  protected
  String getEntityType(String entity)
  {
    if(entity == null || entity.length() == 0) {
      return "";
    }
    int sep = entity.lastIndexOf('.');
    if(sep > 0) {
      return entity.substring(sep + 1).toLowerCase();
    }
    return "";
  }
  
  protected
  void sendBadRequest(HttpServletRequest request, HttpServletResponse response)
    throws ServletException, IOException
  {
    response.sendError(400); // Bad Request
  }
  
  protected
  void sendEntityNotFound(HttpServletRequest request, HttpServletResponse response)
    throws ServletException, IOException
  {
    response.sendError(404); // Not Found
  }
  
  protected
  void sendFile(HttpServletRequest request, HttpServletResponse response, String entityName, String entityType, byte[] content)
    throws ServletException, IOException
  {
    if(content == null) content = new byte[0];
    
    // Send file
    response.setContentType(ExportAs.getContentType(entityType));
    response.addHeader("content-disposition", "attachment; filename=\"" + entityName + "." + entityType +"\"");
    response.setContentLength(content.length);
    
    OutputStream out = response.getOutputStream();
    out.write(content, 0, content.length);
  }
  
  protected
  List<List<Object>> getData(HttpServletRequest request, String entityName, String entityType)
  {
    List<List<Object>> listResult = new ArrayList<List<Object>>();
    
    if(entityName == null || entityName.length() == 0) {
      return listResult;
    }
    
    if(entityName.equalsIgnoreCase("headers")) {
      listResult.add(record("Name", "Value"));
      
      Enumeration<String> enumeration = request.getHeaderNames();
      while(enumeration.hasMoreElements()) {
        String headerName  = enumeration.nextElement();
        String headerValue = request.getHeader(headerName);
        
        listResult.add(record(headerName, headerValue));
      }
    }
    else if(entityName.equalsIgnoreCase("notfound")) {
      return null; // Send Not Found
    }
    
    return listResult;
  }
  
  protected
  List<Object> record(Map<?, ?> map, Object... keys)
  {
    if(keys == null || keys.length == 0 || map == null) {
      return new ArrayList<Object>(0);
    }
    
    List<Object> listResult = new ArrayList<Object>(keys.length);
    for(int i = 0; i < keys.length; i++) {
      listResult.add(map.get(keys[i]));
    }
    
    return listResult;
  }
  
  protected
  List<Object> record(Object... objects)
  {
    if(objects == null || objects.length == 0) {
      return new ArrayList<Object>(0);
    }
    
    List<Object> listResult = new ArrayList<Object>(objects.length);
    for(int i = 0; i < objects.length; i++) {
      listResult.add(objects[i]);
    }
    
    return listResult;
  }
}
