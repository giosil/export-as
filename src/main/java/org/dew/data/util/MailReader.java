package org.dew.data.util;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.PrintStream;

import java.util.Properties;

import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Flags;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.InternetAddress;

public 
class MailReader 
{
  public static
  byte[] readMailBox(String smtp_host, String pop_host, String user, String password, PrintStream psLog)
    throws Exception 
  {
    ByteArrayOutputStream result = new ByteArrayOutputStream();
    
    if(smtp_host == null || smtp_host.length() == 0) {
      return result.toByteArray();
    }
    if(pop_host == null || pop_host.length() == 0) {
      return result.toByteArray();
    }
    if(user == null || user.length() == 0) {
      return result.toByteArray();
    }
    if(psLog == null) {
      psLog = System.out;
    }
    
    String sLastFrom = null;
    
    Properties props = System.getProperties();
    props.put("mail.smtp.host", smtp_host);
    props.put("mail.smtp.auth", "true");
    
    psLog.println("Session.getInstance (mail.smtp.host=" + smtp_host + ")...");
    Session session = Session.getInstance(props, new Authenticator() {
      public PasswordAuthentication getPasswordAuthentication() {
        return new PasswordAuthentication(user, password);
      }
    });
    if(session == null) {
      throw new Exception("Session not available");
    }
    
    psLog.println("session.getStore(\"pop3\")...");
    Store store = session.getStore("pop3");
    if(store == null) {
      throw new Exception("Store not available");
    }
    psLog.println("store.connect(\"" + pop_host + "\",\"" + user + "\",\"" +  password + "\")...");
    store.connect(pop_host, user, password);
    
    psLog.println("store.getDefaultFolder()...");
    Folder defaultFolder = store.getDefaultFolder();
    if(defaultFolder == null) {
      throw new Exception("DefaultFolder not available");
    }
    
    psLog.println("defaultFolder.getFolder(\"INBOX\")...");
    Folder folderInBox = defaultFolder.getFolder("INBOX");
    if(folderInBox == null) {
      throw new Exception("Folder INBOX not available");
    }
    
    byte[] fileContent = null;
    String sSubject = null;
    try {
      psLog.println("folderInBox.open(Folder.READ_WRITE)...");
      folderInBox.open(Folder.READ_WRITE);
      
      psLog.println("folderInBox.getMessages()...");
      Message[] amMessages = folderInBox.getMessages();
      
      if(amMessages == null) {
        amMessages = new Message[0];
        psLog.println("folderInBox.getMessages() -> null");
      }
      else if(amMessages.length == 0) {
        psLog.println("folderInBox.getMessages() -> []");
      }
      else {
        psLog.println("folderInBox.getMessages() -> " + amMessages.length  + " messages");
        
        for(int m = 0; m < amMessages.length; m++) {
          
          Message message = amMessages[m];
          if(message == null) {
            psLog.println("amMessages[" + m + "] is null");
            continue;
          }
          
          InternetAddress[] arrayOfInternetAddress = (InternetAddress[]) message.getFrom();
          if(arrayOfInternetAddress == null || arrayOfInternetAddress.length == 0) {
            psLog.println("amMessages[" + m + "].getFrom() is empty");
            continue;
          }
          
          InternetAddress fromAddress = arrayOfInternetAddress[0];
          String sFrom = null;
          if(fromAddress == null) {
            psLog.println("arrayOfInternetAddress[0].getAddress() -> null");
          }
          else {
            sFrom    = fromAddress.getAddress();
            psLog.println("arrayOfInternetAddress[0].getAddress() -> " + sFrom);
            sLastFrom = sFrom;
          }
          sSubject = message.getSubject();
          psLog.println("message.getSubject() -> " + sSubject);
          
          boolean exceptionDuringDownloadFile = false;
          
          Object content  = message.getContent();
          if(content == null) {
            psLog.println("amMessages[" + m + "].getContent() is null");
            continue;
          }
          if(content instanceof Multipart) {
            psLog.println("amMessages[" + m + "].getContent() is Multipart");
            Multipart multipart = (Multipart) content;
            
            for(int i = 0; i < multipart.getCount(); i++) {
              
              BodyPart bodyPart = multipart.getBodyPart(i);
              if(bodyPart == null) {
                psLog.println("multipart.getBodyPart(" + i + ") -> null");
                continue;
              }
              
              String sDisposition = bodyPart.getDisposition();
              psLog.println("multipart.getBodyPart(" + i + ").getDisposition() -> " + sDisposition);
//              if(sDisposition == null || !sDisposition.equalsIgnoreCase(Part.ATTACHMENT)) {
//                psLog.println("multipart.getBodyPart(" + i + ").getDisposition() is not " + Part.ATTACHMENT);
//                continue;
//              }
              
              String sFileName = bodyPart.getFileName();
              psLog.println("multipart.getBodyPart(" + i + ").getFileName() -> " + sFileName);
              if(sFileName == null || sFileName.length() == 0) {
                psLog.println("multipart.getBodyPart(" + i + ").getFileName() is empty");
                continue;
              }
              
              InputStream is = bodyPart.getInputStream();
              if(is == null) {
                psLog.println("multipart.getBodyPart(" + i + ").getInputStream() is null");
                continue;
              }
              
              psLog.println("  download " + sFileName + "...");
              try {
                byte[] buffer = new byte[4096];
                int bytesRead;
                while((bytesRead = is.read(buffer)) != -1) {
                  result.write(buffer, 0, bytesRead);
                }
              }
              catch(Exception ex) {
                psLog.println("Exception during download " + sFileName + ": " + ex);
                exceptionDuringDownloadFile = true;
              }
              finally {
                if(result != null) try { result.close(); } catch(Exception ex) {}
              }
              
              fileContent = result.toByteArray();
              if(fileContent != null && fileContent.length > 0) {
                return fileContent;
              }
            }
          }
          else if(content instanceof String) {
            psLog.println("amMessages[" + m + "].getContent() = \"" + content + "\"");
          }
          
          if(!exceptionDuringDownloadFile) {
            psLog.println("message.setFlag(Flags.Flag.DELETED, true)...");
            message.setFlag(Flags.Flag.DELETED, true);
          }
        }
      }
    }
    finally {
      psLog.println("folderInBox.close(true)...");
      folderInBox.close(true);
      
      psLog.println("store.close()...");
      store.close();
    }
    
    if(sLastFrom == null || sLastFrom.length() == 0) {
      sLastFrom = "";
    }
    else {
      sLastFrom = "(" + sLastFrom + ")";
    }
    return result.toByteArray();
  }
}
