package com.theta.common.app.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.security.MessageDigest;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Enumeration;
import java.util.List;

import javax.servlet.http.HttpServletRequest;

import jxl.format.UnderlineStyle;
import jxl.write.DateFormat;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;

import org.apache.log4j.Logger;
import org.dom4j.Node;

import com.theta.common.exception.MspplusException;


public class BaseUtil {

  private static Logger log = Logger.getLogger(BaseUtil.class);

  private static long   oneDay = 24 * 3600000;
  
  
  /*
   * 预定义的一些字体和格式，同一个Excel中最好不要有太多格式   
   */   
 static WritableFont headerFont = new WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD, 
    false, UnderlineStyle.NO_UNDERLINE, jxl.format.Colour.BLUE);    
 public static WritableCellFormat headerFormat = new WritableCellFormat (headerFont);    
   
   
 static WritableFont detFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.NO_BOLD, 
    false, UnderlineStyle.NO_UNDERLINE, jxl.format.Colour.BLACK);    
 public static WritableCellFormat detFormat = new WritableCellFormat (detFont);    
   
 static NumberFormat nf=new NumberFormat("0"); //用于Number的格式   
 public static WritableCellFormat numberFormat = new WritableCellFormat (detFont, nf);  

   
 static DateFormat df=new DateFormat("yyyy-MM-dd HH:mm:ss");//用于日期的   
 public static WritableCellFormat dateFormat = new WritableCellFormat(detFont, df);
 
	 
  /**
   * 取得当前时间的字符串 格式为类似: 20051102120356125
   * 
   * @return String 返回当前时间的字符串
   */
  public static String getCurrentTime() {
    SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMddHHmmssSSS");
    return formatter.format(new Date());
  }

  /**
   * 取得当前时间的字符串 格式为自定义
   * 
   * @param format - 返回自定义的日期格式
   * @return String 返回当前时间的字符串
   */
  public static String getCurrentTime(String format) {
    if (format != null && !"".equals(format.trim())) {
	      SimpleDateFormat formatter = new SimpleDateFormat(format.trim());
	      return formatter.format(new Date());
    }
    return "";
  }

  /**
   * 取得时间的字符串 格式为自定义
   * 
   * @param format - 返回自定义的日期格式
   * @param time - 返回要转换的时间
   * @return String 返回当前时间的字符串
   */
  public static String getMyTime(long time, String format) {
    if (time > 0 && format != null && !"".equals(format.trim())) {
      SimpleDateFormat formatter = new SimpleDateFormat(format.trim());
      return formatter.format(new Date(time));
    }
    return "";
  }

  /**
   * 取得时间的字符串 格式为自定义
   * 
   * @param format - 返回自定义的日期格式
   * @param time - 返回要转换的时间
   * @return String 返回当前时间的字符串
   */
  public static String getMyTime(Date time, String format) {
    String rtnTime = "";
    if (time != null && format != null && !"".equals(format.trim())) {
      SimpleDateFormat formatter = new SimpleDateFormat(format.trim());
      rtnTime = formatter.format(time);
    }
    return rtnTime;
  }

  /**
   * 取得时间的字符串 格式为自定义
   * 
   * @param format - 返回自定义的日期格式
   * @param time - 返回要转换的时间
   * @return String 返回当前时间的字符串
   */
  public static String getMyTime(Calendar time, String format) {
    String rtnTime = "";
    if (time != null && format != null && !"".equals(format.trim())) {
      SimpleDateFormat formatter = new SimpleDateFormat(format.trim());
      rtnTime = formatter.format(time.getTime());
    }
    return rtnTime;
  }

  
  /**
   * 字符串转换成整型
   * 
   * @param str - 字符串.
   * @return int - 返回整型.
   * @throws CommonException - 失败时抛出异常。
   */
  public static int StrToInt(String str) throws MspplusException {
    int nResult = 0;
    try {
      nResult = new Integer(str).intValue();
    }
    catch (Exception ae) {
      log.error("数字转换错误!Str:"+str, ae);
      throw new MspplusException("数字转换错误!Str:"+str);
    }
    return nResult;
  }

  /**
   * 将数据库中取出的 String [yyyy-mm-dd hh:mm:ss]转为Date 字符串类型 [yyyy-mm-dd hh:ss:mm]
   * [yyyy/mm/dd hh:ss:mm] [yyyy-mm-dd]
   * 
   * @param strDate the parseDate
   * @throws CEPMSException
   * @see java.util.Date
   * @see java.util.Calendar
   * @return Date
   */
  public static Date parseDate(String strDate) throws MspplusException {
    Date date = null;

    try { // try 1
      if (strDate != null && !strDate.trim().equals("")) { // if 1
        String y = strDate.substring(0, 4);
        String m = strDate.substring(5, 7);
        String d = strDate.substring(8, 10);
        String h = "0";
        String mi = "0";
        String se = "0";

        if (strDate.length() > 11) {
          h = strDate.substring(11, 13);
          mi = strDate.substring(14, 16);
          se = strDate.substring(17, 19);
        }

        int year = Integer.parseInt(y);
        int month = Integer.parseInt(m) - 1;
        int day = Integer.parseInt(d);
        int hour = Integer.parseInt(h);
        int minutes = Integer.parseInt(mi);
        int seconds = Integer.parseInt(se);

        Calendar cal = Calendar.getInstance();
        cal.set(year, month, day, hour, minutes, seconds);
        cal.set(Calendar.MILLISECOND, 0);
        date = cal.getTime();
      } // end if 1
    }catch (NumberFormatException nfe) { // catch 1
      log.error("转为Date错误：error is:" , nfe);
      throw new MspplusException("输入日期格式错误:" + nfe);
    } // end try 1

    return date;
  }

  /**
   * 将数据库中取出的 String [yyyy-mm-dd hh:ss:mm]转为Date 字符串类型 [yyyy-mm-dd hh:ss:mm]
   * [yyyy/mm/dd hh:ss:mm] [yyyy-mm-dd]
   * 
   * @param strDate the parseDate
   * @param type 转换的格式 yyyy-MM-dd 等等
   * @throws CEPMSException
   * @see java.util.Date
   * @return Date
   */
  public static Date parseDate(String strDate, String type)
      throws MspplusException {
    
    if (strDate == null || strDate.equals("")) {
      return null;
    }

    Date dt = null;
    SimpleDateFormat format = null;

    if (type == null || type.equals("")) {
      format = new SimpleDateFormat("yyyy-MM-dd hh:ss:mm");
    }
    else {
      format = new SimpleDateFormat(type);
    }
    if(strDate.trim().length()==10)
    {
    	format = new SimpleDateFormat("yyyy-MM-dd");
    }else if(strDate.trim().length()==7){
    	format = new SimpleDateFormat("yyyy-MM");
    }
    
    try {
      dt = format.parse(strDate);
    }
    catch (ParseException e) {
      log.error("解析日期错误, strDate:"+ strDate +" type:"+ type, e);
      throw new MspplusException("请按格式["+type+"]输入日期!");
    }
    return dt;
  }

  public static Date parseDateExt(String strDate, String type, String errorMsg) throws MspplusException {

		if (strDate == null || strDate.equals("")) {
		  return null;
		}
		
		Date dt = null;
		SimpleDateFormat format = null;
		
		if (type == null || type.equals("")) {
		  format = new SimpleDateFormat("yyyy-MM-dd hh:ss:mm");
		}
		else {
		  format = new SimpleDateFormat(type);
		}
		try {
		  dt = format.parse(strDate);
		}
		catch (ParseException e) {
		  log.error("解析日期错误, strDate:"+ strDate +" type:"+ type, e);
		  throw new MspplusException(errorMsg+"请按格式["+type+"]输入日期!");
		}
		return dt;
}
  
  /**
   * 将时间型DATE 转为 字符串类型 [yyyy-MM-dd hh:ss:mm] [yyyy/MM/dd hh:ss:mm] [yyyy-MM-dd]
   * 
   * @param Date the datetime
   * @param String 转换的格式 yyyy-MM-dd 等等
   * @throws CEPMSException
   * @see java.util.Date
   * @see java.text.SimpleDateFormat
   * @return String
   */
  public static String parseDate(Date datetime, String formatType)
      throws MspplusException {

    String strDate = "";
    try {
      if (datetime != null && formatType != null) {
        SimpleDateFormat format = new SimpleDateFormat(formatType);
        strDate = format.format(datetime);
      }
    }
    catch (Exception e) {
      log.error("转为String错误：error is:" , e);
      throw new MspplusException("转为Date错误：error is:" + e);
    }
    return strDate;
  }
  
  public static String parseString(String dateStr) {
	  String strDate = "";
	  Date destDate = null;
	  try {
		  if (dateStr != null && dateStr != null) {
			  SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd hh:ss:mm");
			  destDate = format.parse(dateStr);
			  destDate.setMonth(destDate.getMonth()+1);
			  strDate = format.format(destDate);
			  }
		  } catch (Exception e) {
			  log.error("转为String错误：error is:" , e);
			  }
		  return strDate;
  }


  /**
   * 返回大于一天日期字符串。
   * @param str
   * @param format
   * @return
 * @throws MspplusException 
   */
  public static String getDateLargeOneDay(String strDate , String format) throws MspplusException{
	  
	   Date date = parseDate(strDate, format);
	   
	   
	   Date date2 = new Date(date.getTime()+oneDay);
	   
	  return  parseDate(date2, format);
	 
  }
  
  /**
   * 该方法完成md5加密算法,返回加密后的字符串
   * 
   * @param input the String
   * @return a<code>String</code>specifying the encrypted character string
   * @throws CEPMSException
   */
  public static String md5(String input) throws MspplusException {

    MessageDigest md5 = null;

    try {
      md5 = MessageDigest.getInstance("md5");
    }
    catch (java.security.NoSuchAlgorithmException nsae) {
      throw new MspplusException(nsae.getMessage());
    }

    char[] charArray = input.toCharArray();
    byte[] byteArray = new byte[charArray.length];
    for (int i = 0; i < charArray.length; i++)
      byteArray[i] = (byte)charArray[i];

    byte[] md5Bytes = md5.digest(byteArray);
    StringBuffer hexValue = new StringBuffer();

    for (int i = 0; i < md5Bytes.length; i++) {
      int val = ((int)md5Bytes[i]) & 0xff;
      if (val < 16)
        hexValue.append("0");
      hexValue.append(Integer.toHexString(val));
    }
    return hexValue.toString().toUpperCase();
  }
  /**
   * 根据用户名和密码进行MD5加密
   * @param userName
   * @param password
   * @return
   * @throws MspplusException
   */
  public static String md5FromUsernameAndPassword(String userName,String password) throws MspplusException {
	  
	 String input = password+"@"+userName;
	 return md5(input);
  }

  /**
   * 从文件中读取二进制数据，并返回
   * 
   * @param file
   * @return byte[] - 返回读入的二进制数据
   * @throws CEPMSException
   */
  public static byte[] getBuffer(File file) throws MspplusException {
    byte[] buffer = null;

    try {
      if (file != null && file.length() > 0) {
        int length = (int)file.length();
        buffer = new byte[length];

        FileInputStream fis = new FileInputStream(file);

        while (fis.available() != 0) {
          fis.read(buffer);
        }
      }
      else {
        log.info("传入的文件为null或文件的长度为0。");
      }
    }
    catch (FileNotFoundException fne) {
      log.error("从文件中取得byte[]异常, exception is:" + file);
      throw new MspplusException(fne.getMessage());
    }
    catch (IOException ioe) {
      log.error("从文件中取得byte[]异常, exception is:" + file);
      throw new MspplusException(ioe.getMessage());
    }

    return buffer;
  }

  /**
   * 返回当前的url
   * 
   * @param HttpServletRequest
   * @param List --要忽略掉
   * @return String url地址
   */
  @SuppressWarnings("unchecked")
public static String getUrl(List list, HttpServletRequest request) {

    String url = "";
    Enumeration param = request.getParameterNames();// 得到所有参数名
    while (param.hasMoreElements()) {
      String pname = param.nextElement().toString().trim();
      boolean flag = true;
      if (list != null && !list.isEmpty()) {
        for (int i = 0; i < list.size(); i++) {
          String button = (String)list.get(i);
          if (pname.equals(button)) {
            flag = false;
          }
        }

        if (flag) {
          url += pname + "=" + request.getParameter(pname).trim() + "&";
        }
      }
    }
    if (url.endsWith("&")) {
      url = url.substring(0, url.lastIndexOf("&"));
    }
    String urls = request.getRequestURL().toString();

    if (!url.equals("")) {
      urls = request.getRequestURL() + "?" + url;
    }
    else {
      urls = request.getRequestURL().toString();
    }
    return urls;
  }
  
  /**
   * 将long型转化成字符串的ip[102.23.123.12]
   * @param lip
   * @return
   */
  public static String long2Ip(long lip) {
    String sip = "";
    
    String one = String.valueOf(lip & 0xff);
    String two = String.valueOf((lip >> 8) & 0xff);
    String three = String.valueOf((lip >> 16) & 0xff);
    String four = String.valueOf((lip >> 24) & 0xff);

    sip = one + "." + two + "." + three + "." + four;
    
    return sip;
  }
  
  /**
   * 将时间的时、分、秒都设为0
   * @param date
   * @return
   */
  public static Date setMinTime(Date date) {
    if (date == null) return null;
    
    Calendar cal = Calendar.getInstance();
    cal.setTime(date);
    cal.set(Calendar.HOUR_OF_DAY, 0);
    cal.set(Calendar.MINUTE, 0);
    cal.set(Calendar.SECOND, 0);
    cal.set(Calendar.MILLISECOND, 0);
    
    return cal.getTime();
  }
  
  /**
   * 将时间的时、分、秒都设为最大
   * @param date
   * @return
   */
  public static Date setMaxTime(Date date) {
    if (date == null) return null;
    
    Calendar cal = Calendar.getInstance();
    cal.setTime(date);
    cal.set(Calendar.HOUR_OF_DAY, 23);
    cal.set(Calendar.MINUTE, 59);
    cal.set(Calendar.SECOND, 59);
    cal.set(Calendar.MILLISECOND, 999);
    
    return cal.getTime();  
  }
  
	

  /**
   * 
   * @param ip_address
   * @return
   */
  public static String getIpToLongStr(String ip_address){
  	

		int index1 =  ip_address.indexOf(".");
		
		String temp = ip_address.substring(0, index1);
		
		long ip = 0;
		
		ip += (Long.parseLong(temp)&0xFF) ;
		
		
	    int index2 = ip_address.indexOf(".", index1+1);
	    temp = ip_address.substring(index1+1, index2);
		ip += ((Long.parseLong(temp)<<8)&0xFF00);
		
		 
		int index3 = ip_address.indexOf(".", index2+1); 
		temp = ip_address.substring(index2+1, index3);
		ip += ((Long.parseLong(temp)<<16)&0xFF0000);
		 
		
		temp = ip_address.substring(index3+1);
		ip += ((Long.parseLong(temp)<<32)&0xFF000000);
		
		return String.valueOf(ip);

  	}

  	/**
  	 * 
  	 * @param node
  	 * @param xpath
  	 * @return
  	 */
  	public static String getXpathValue(Node node , String xpath){
  		
  		if(node==null){
  			return null;
  		}
  		
  		Node temp = node.selectSingleNode(xpath);
  		if(temp==null){
  			return null;
  		}
  		return temp.getText();
  	}

	// 十六进制转换为字符串
	public static String toStringHex(String s) {
		s=s.replaceAll("0x","");
		byte[] baKeyword = new byte[s.length() / 2];
		for (int i = 0; i < baKeyword.length; i++) {
			try {
				baKeyword[i] = (byte) (0xff & Integer.parseInt(s.substring(i * 2, i * 2 + 2), 16));
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		try {
			s = new String(baKeyword, "utf-8");// UTF-16le:Not
		} catch (Exception e1) {
			e1.printStackTrace();
		}
		return s;
	}
  	
  	
}
