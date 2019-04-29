package com.theta.report.ver1;

import java.io.IOException;
import java.io.Writer;

import jxl.format.Alignment;
import jxl.write.DateFormat;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFormat;
import jxl.write.WriteException;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import org.apache.log4j.Logger;

public class FileExportInsertAndUpdateSqlTextUtil {

	public static final int InsertTxtType = 10;
	public static final int UpdateTxtType = 11;
	public static final int PageSize = 60000;
	protected static Logger logger =  Logger.getLogger(FileExportInsertAndUpdateSqlTextUtil.class) ;

	protected static	NumberFormat doubleNf = new NumberFormat("#.##");
	protected static	NumberFormat intNf = new NumberFormat("#");
	
	protected static  	WritableCellFormat doubleCf = new WritableCellFormat(doubleNf);
	protected static  	WritableCellFormat intCf = new WritableCellFormat(intNf);
	
	private static WritableCellFormat dateCf = null;

	public static WritableCellFormat getDoubleFormat(){
	
		return doubleCf;
	}
	
	public static WritableCellFormat getIntFormat(){	
		
		return intCf;
	}
	
	public static WritableCellFormat getDateCellFormat(){

		if(dateCf==null){
			DateFormat df = new DateFormat("yyyy-MM-dd hh:mm:ss");
			dateCf = new WritableCellFormat(df); 
		}
		try {
			dateCf.setAlignment(Alignment.LEFT);
		} catch (WriteException e) {
			logger.error(e);
		}
		return dateCf;
	
	}

	/**
	 * 导出表 sys_host_label 的 Insert的SQL语句
	 *dataIndex[0]为序列，故应从dataIndex[1]开始读取，
	 * @param writer
	 * @param fileName
	 * @param header
	 * @param dataIndex
	 * @param array
	 * @param hiddens
	 */
	public static void exportSiteClassInsertSqlTxt(Writer writer, String fileName,
			String[] header, String[] dataIndex, JSONArray array,boolean[] hiddens) {
	
		if(logger.isDebugEnabled()){
			logger.debug("exportTxt be called !");			
		}

		if(header==null){
			return ;
		}
		StringBuffer sb = new StringBuffer();
		
		 
		for(int i = 0; i<array.size(); i++){			 
			boolean first = true;
			sb.append("insert into sys_host_label (");
			for(int j=1; j<dataIndex.length; j++){
				if(!hiddens[j]){
					 if(first){
						 first = false;	
					 }else{
						 
						 sb.append(",");
					 } 
					 sb.append(dataIndex[j]);
				}		 
			}
		
			sb.append(") values (");
			
			 JSONObject ja = (JSONObject) array.get(i);
			 if(ja==null){
				 continue;
			 }
			 first = true;
			 
 			 int iw=1;
			 for(int index=1; index < dataIndex.length; index++){
				 String key =  dataIndex[index];
				 String str = ja.getString(key);
				 if(!hiddens[iw]){
				
					 if(first){
						 first = false;	
					 }else{
						 sb.append(",");
					 }
					 sb.append("'");
					 sb.append(str);
					 sb.append("'");
				 } 
				 iw=iw+1;	
			 }
 
			 sb.append(");\r\n");
		 }
		 
		String re= sb.toString();
	 
		try {
			writer.write(re);
		} catch (IOException e) {
			 logger.error(e);
		}
	}
	/**
	 * 导出表 sys_host_label 的Update的SQL语句
	 * dataIndex[0]为序列，故应从dataIndex[1]开始读取，
	 * 将dataIndex[1]作为修改的where语句，因此从dataIndex[2]开始读取
	 * 
	 * @param writer
	 * @param fileName
	 * @param header
	 * @param dataIndex
	 * @param array
	 * @param hiddens
	 */
	public static void exportSiteClassUpdateSqlTxt(Writer writer, String fileName,
			String[] header, String[] dataIndex, JSONArray array,boolean[] hiddens) {
	
		if(logger.isDebugEnabled()){
			logger.debug("exportTxt be called !");			
		}

		if(header==null){
			return ;
		}
		StringBuffer sb = new StringBuffer();
		
		 
		for(int i = 0; i<array.size(); i++){			 
			boolean first = true;
			
			 JSONObject ja = (JSONObject) array.get(i);
			 
			 if(ja==null){
				 continue;
			 }
			 first = true;
			 
			 sb.append("update sys_host_label set ");
			 String site_type = ja.getString("site_type");
			 
			 String host = ja.getString("host");
			 sb.append("site_type =");
			 sb.append(site_type);
			 sb.append(" where host='");
			 sb.append(host);
			 sb.append("';\r\n");
			 
// 			 int iw=2;
//			 for(int index=2; index < dataIndex.length; index++){
//				 String key = dataIndex[index];
//				 String str = ja.getString(key);
//				
//				 if(!hiddens[iw]){
//				
//					 if(first){
//						 first = false;	
//					 }else{
//						 sb.append(",");
//					 } 
//					 sb.append(key);
//					 sb.append("='");
//					 sb.append(str);
//					 sb.append("'");
//				 } 
//				 iw=iw+1;	
//			 }
// 			 String arrStr[] = new String[3];
//			 arrStr[0] = "host";
//			 arrStr[1] = "site_type";
			 
//			 for(int index=1; index < arrStr.length; index++){
//				 String key = arrStr[index];
//				 String str = ja.getString(key);
//				 
//				 if(!hiddens[iw]){
//					 
//					 if(first){
//						 first = false;	
//					 }else{
//						 sb.append(",");
//					 } 
//					 sb.append(key);
//					 sb.append("='");
//					 sb.append(str);
//					 sb.append("'");
//				 } 
//				 iw=iw+1;	
//			 }
// 
////			 String  key1 = dataIndex[1];
//			 String  key1 = arrStr[0];
//			 String value1 = ja.getString(key1);
//			 sb.append(" where ");
//			 sb.append(key1);
//			 sb.append("='");
//			 sb.append(value1);
//			 sb.append("';\r\n");
		 }
		 
		String re= sb.toString();
	 
		try {
			writer.write(re);
		} catch (IOException e) {
			 logger.error(e);
		}
	}
	
}
