package com.theta.report.ver1;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Writer;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.write.DateFormat;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import org.apache.log4j.Logger;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;

import com.theta.common.app.util.BaseUtil;
import com.theta.common.app.util.PdfGenerator;
import com.theta.report.ver1.dim1.model.data.FieldInfo;
import com.theta.report.ver1.jiekou.IHeader;
import com.theta.report.ver1.jiekou.data.IData;

public class FileExportUtil {

	public static final int ExcelType = 0;
	public static final int TxtType = 1;
	public static final int HtmlType = 2;
	public static final int PdfType = 3;
	public static final int XmlType = 4;
	public static final int TxtTypeNoIndex = 5;
	public static final int ImeiInsertSql = 6;
	public static final int ImsiProperties = 7;
	public static final int MODULELISTSql = 8;
	public static final int PageSize = 60000;

	private static Set<String> doubleToStr = new HashSet<String>();
	static {
		doubleToStr.add("msisdn");
		doubleToStr.add("imsi");
		doubleToStr.add("eci");
		doubleToStr.add("old_eci");
		doubleToStr.add("lac");
		doubleToStr.add("old_lac");
		doubleToStr.add("tac");
		doubleToStr.add("old_tac");
	}

	protected static Logger logger = Logger.getLogger(FileExportUtil.class);

	protected static NumberFormat doubleNf = new NumberFormat("0.00");
	protected static NumberFormat intNf = new NumberFormat("0");

	protected static WritableCellFormat doubleCf = new WritableCellFormat(
			doubleNf);
	protected static WritableCellFormat intCf = new WritableCellFormat(intNf);

	private static WritableCellFormat dateCf = null;

	public static WritableCellFormat getDoubleFormat() {

		return doubleCf;
	}

	public static WritableCellFormat getIntFormat() {

		return intCf;
	}

	public static WritableCellFormat getDateCellFormat() {

		if (dateCf == null) {
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
	 * jxl.write.NumberFormat nf = new jxl.write.NumberFormat("#.##");
	 * jxl.write.WritableCellFormat wcfN = new jxl.write.WritableCellFormat(nf);
	 * jxl.write.Number labelNF = new jxl.write.Number(0, 2, 3.1415926, wcfN);
	 * sheet.addCell(labelNF);
	 * 
	 * Description: excel文件导出
	 * @author lenovo
	 * @param out 文件输出流
	 * @param sheetName 工作表名称
	 * @param header 表头
	 * @param dataIndex 与查询结果JSONObject的key对应的名称
	 * @param array 查询结果信息
	 * @param hiddens 列是否显示
	 * @exception IOException,RowsExceededException,WriteException
	 */
	public static void exportExcel(OutputStream out, String sheetName,
			String[] header, String[] dataIndex, JSONArray array,
			boolean[] hiddens) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportExcel be called !");
		}

		// 一些临时变量，用于写到excel中
		Label l = null;

		// add head
		int column = 0;
		int row = 0;

		WritableWorkbook workbook = null;

		try {

			workbook = Workbook.createWorkbook(out);

			WritableSheet sheet = workbook.createSheet(sheetName, 1);

			for (int i = 0; i < header.length; i++) {
				if (!hiddens[i]) {
					l = new Label(column++, row, header[i],
							BaseUtil.headerFormat);
					sheet.addCell(l);
				}
			}
			for (int i = 0; i < array.size(); i++) {
				column = 0;
				row = row + 1;

				if (i != 0 && i % PageSize == 0) {
					for (int is = 0; is < sheet.getColumns(); is++) {

						sheet.setColumnView(is, header.length);

					}
					int bh = i / PageSize;
					sheet = workbook.createSheet(String.valueOf(bh), bh);
					row = 0;
				}

				JSONObject ja = (JSONObject) array.get(i);
				if (ja == null) {
					continue;
				}
				int iw = 0;
				for (int index = 0; index < header.length; index++) {

					String key = dataIndex[index];
					String str = ja.getString(key);
					if (!hiddens[iw]) {

						if (str == null) {
							str = "";
						}
						l = new Label(column++, row, str, BaseUtil.detFormat);
						sheet.addCell(l);
					}
					iw = iw + 1;
				}
			}

			for (int i = 0; i < sheet.getColumns(); i++) {

				sheet.setColumnView(i, header.length);

			}

			workbook.write();
			workbook.close();

		} catch (Exception e) {
			logger.error(e);
		}
	}

	/**
	 * jxl.write.NumberFormat nf = new jxl.write.NumberFormat("#.##");
	 * jxl.write.WritableCellFormat wcfN = new jxl.write.WritableCellFormat(nf);
	 * jxl.write.Number labelNF = new jxl.write.Number(0, 2, 3.1415926, wcfN);
	 * sheet.addCell(labelNF);
	 * 
	 * Description: excel文件导出
	 * @author lenovo
	 * @param out 文件输出流
	 * @param sheetName 工作表名
	 * @param header 显示列的各种属性信息
	 * @param records 查询结果信息
	 * @param starttime 查询数据的开始时间
	 * @param endtime 查询数据的结束时间
	 * @exception IOException,RowsExceededException,WriteException
	 */
	public static void exportExcel(OutputStream out, String sheetName,
			IHeader[] header, JSONArray records, String starttime,
			String endtime) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportExcel be called !");
		}

		// 一些临时变量，用于写到excel中
		Label l = null;

		// add head
		int column = 0;
		int row = 0;

		WritableWorkbook workbook = null;
		try {

			workbook = Workbook.createWorkbook(out);

			WritableSheet sheet1 = workbook.createSheet("查询时间", 0);

			String title = sheetName.split("2")[0];
			Label queryTime = new Label(0, 0, title, BaseUtil.headerFormat);
			sheet1.addCell(queryTime);

			String starttimes = "开始时间:" + starttime;
			Label queryTime1 = new Label(0, 1, starttimes);
			sheet1.addCell(queryTime1);

			String endtimes = "结束时间:" + endtime;
			Label queryTime2 = new Label(0, 2, endtimes);
			sheet1.addCell(queryTime2);

			WritableSheet sheet = workbook.createSheet(sheetName, 1);

			for (int i = 0; i < header.length; i++) {
				if (header[i].isExport()) {
					l = new Label(column++, row, header[i].getHeaderName(),
							BaseUtil.headerFormat);
					sheet.addCell(l);
				}
			}
			for (int i = 0; i < records.size(); i++) {
				column = 0;
				row = row + 1;
				if (i != 0 && i % PageSize == 0) {
					int colIndex = 0;
					for (int is = 0; is < header.length; is++) {
						if (!header[is].isExport()) {
							continue;
						}
						if (header[is].isSetWidth()) {
							sheet.setColumnView(colIndex,
									header[is].getHeaderWidth() / 5);
						} else {
							sheet.setColumnView(colIndex, header[is]
									.getHeaderName().length() * 2);
						}
						colIndex++;
					}
					int bh = i / PageSize;
					sheet = workbook.createSheet(String.valueOf(bh), bh);
					row = 0;
				}

				JSONObject ja = (JSONObject) records.get(i);
				if (ja == null) {
					continue;
				}
				int headIndex = 0;
				for (headIndex = 0; headIndex < header.length; headIndex++) {
					IHeader thead = header[headIndex];

					String key = thead.getDataIndex();
					boolean bError = false;
					if (thead.isExport()) {

						try {
							if ((thead.isDouble() || thead.isInt())
									&& ja.getDouble(key) > Integer.MAX_VALUE - 1) {
								thead.setHeadType(IHeader.Type_String);
							}
						} catch (Exception e) {
							// TODO Auto-generated catch block
							logger.error("The value is not a Double or int Value : "
									+ ja.getString(key) + " " + e);
						}

						if (thead.isDouble()) {
							try {
								WritableCellFormat df = getDoubleFormat();
								double d = ja.getDouble(key);
								Number nb = new jxl.write.Number(column, row,
										d, df);
								sheet.addCell(nb);
								column++;
							} catch (Exception ex) {
								thead.setHeadType(IHeader.Type_String);
								bError = true;
								logger.error(ex);
							}

						} else if (thead.isInt()) {
							try {
								WritableCellFormat df = getIntFormat();
								int d = ja.getInt(key);
								Number nb = new jxl.write.Number(column, row,
										d, df);
								sheet.addCell(nb);
								column++;
							} catch (Exception ex) {
								thead.setHeadType(IHeader.Type_String);
								bError = true;
								logger.error("key:" + key, ex);
							}
						} else if (thead.isLong()) {
							try {
								// 长整形时按字符串输出
								long d = ja.getLong(key);
								l = new Label(column++, row, d + "",
										BaseUtil.detFormat);
								sheet.addCell(l);

							} catch (Exception ex) {
								thead.setHeadType(IHeader.Type_String);
								bError = true;
								logger.error("key:" + key, ex);
							}

						} else {
							String str = ja.getString(key);
							if (doubleToStr.contains(thead.getDataIndex())) {
								// 数据库msisdn和imsi用的number型时为科学计数法，转换为正确的字符串
								if (!"".equals(str)) {
									java.text.NumberFormat nf = java.text.NumberFormat
											.getInstance();
									nf.setGroupingUsed(false);
									str = nf.format(Double.valueOf(str));
								}
							}

							if (str == null) {
								str = "";
							}
							l = new Label(column++, row, str,
									BaseUtil.detFormat);
							sheet.addCell(l);
						}

						if (bError) {
							String str = ja.getString(key);
							if (str == null) {
								str = "";
							}
							l = new Label(column++, row, str,
									BaseUtil.detFormat);
							sheet.addCell(l);
						}
					}
				}
			}

			int colIndex = 0;
			for (int i = 0; i < header.length; i++) {
				if (!header[i].isExport()) {
					continue;
				}
				if (header[i].isSetWidth()) {
					sheet.setColumnView(colIndex,
							header[i].getHeaderWidth() / 5);
				} else {
					sheet.setColumnView(colIndex, header[i].getHeaderName()
							.length() * 2);
				}
				colIndex++;
			}

			workbook.write();
			workbook.close();

		} catch (Exception e) {
			logger.error(e);
		}
	}
	
	/**
	 * 
	 * Description: excel文件导出
	 * @author lenovo
	 * @param out 文件输出流
	 * @param sheetName 工作表名
	 * @param header 显示列的各种属性信息
	 * @param records 查询结果信息
	 * @param starttime 查询数据的开始时间
	 * @param endtime 查询数据的结束时间
	 * @param groupHeaders 表格列所占大小合并信息
	 * @exception IOException,RowsExceededException,WriteException
	 */
	public static void exportExcel(OutputStream out, String sheetName,
			IHeader[] header, JSONArray records, String starttime,
			String endtime,String groupHeaders) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportExcel be called !");
		}

		// 一些临时变量，用于写到excel中
		Label l = null;

		// add head
		int column = 0;
		int row = 0;

		WritableWorkbook workbook = null;
		try {

			workbook = Workbook.createWorkbook(out);

			WritableSheet sheet1 = workbook.createSheet("查询时间", 0);

			String title = sheetName.split("2")[0];
			Label queryTime = new Label(0, 0, title, BaseUtil.headerFormat);
			sheet1.addCell(queryTime);

			String starttimes = "开始时间:" + starttime;
			Label queryTime1 = new Label(0, 1, starttimes);
			sheet1.addCell(queryTime1);

			String endtimes = "结束时间:" + endtime;
			Label queryTime2 = new Label(0, 2, endtimes);
			sheet1.addCell(queryTime2);

			WritableSheet sheet = workbook.createSheet(sheetName, 1);
			
			String[] groupHeaders1 = groupHeaders.split(";");
			for (int i = 0; i < groupHeaders1.length; i++) {
				String[] headers1 = groupHeaders1[i].split(",");
				int startIndex = 0;
				int endIndex = 0;
				logger.info("headers length:"+headers1.length);
				for (int j = 0; j < headers1.length; j++) {
					String[] headerName = headers1[j].split(":");
					String name = headerName[0];
					int col = Integer.parseInt(headerName[1]);
					l = new Label(startIndex, i, name);
					sheet.addCell(l);
					
					if(col>1){
						endIndex = startIndex+col-1;
						sheet.mergeCells(startIndex, i, endIndex, i);
						startIndex = endIndex+1;
					}else{
						startIndex = j+endIndex+1;
					}
				}
			}
			
			row = groupHeaders1.length;

			for (int i = 0; i < header.length; i++) {
				if (header[i].isExport()) {
					l = new Label(column++, row, header[i].getHeaderName());
					sheet.addCell(l);
				}
			}
			for (int i = 0; i < records.size(); i++) {
				column = 0;
				row = row + 1;
				if (i != 0 && i % PageSize == 0) {
					int colIndex = 0;
					for (int is = 0; is < header.length; is++) {
						if (!header[is].isExport()) {
							continue;
						}
						if (header[is].isSetWidth()) {
							sheet.setColumnView(colIndex,
									header[is].getHeaderWidth() / 5);
						} else {
							sheet.setColumnView(colIndex, header[is]
									.getHeaderName().length() * 2);
						}
						colIndex++;
					}
					int bh = i / PageSize;
					sheet = workbook.createSheet(String.valueOf(bh), bh);
					row = 0;
				}

				JSONObject ja = (JSONObject) records.get(i);
				if (ja == null) {
					continue;
				}
				int headIndex = 0;
				for (headIndex = 0; headIndex < header.length; headIndex++) {
					IHeader thead = header[headIndex];

					String key = thead.getDataIndex();
					boolean bError = false;
					if (thead.isExport()) {

						try {
							if ((thead.isDouble() || thead.isInt())
									&& ja.getDouble(key) > Integer.MAX_VALUE - 1) {
								thead.setHeadType(IHeader.Type_String);
							}
						} catch (Exception e) {
							// TODO Auto-generated catch block
							logger.error("The value is not a Double or int Value : "
									+ ja.getString(key) + " " + e);
						}

						if (thead.isDouble()) {
							try {
								WritableCellFormat df = getDoubleFormat();
								double d = ja.getDouble(key);
								Number nb = new jxl.write.Number(column, row,
										d, df);
								sheet.addCell(nb);
								column++;
							} catch (Exception ex) {
								thead.setHeadType(IHeader.Type_String);
								bError = true;
								logger.error(ex);
							}

						} else if (thead.isInt()) {
							try {
								WritableCellFormat df = getIntFormat();
								int d = ja.getInt(key);
								Number nb = new jxl.write.Number(column, row,
										d, df);
								sheet.addCell(nb);
								column++;
							} catch (Exception ex) {
								thead.setHeadType(IHeader.Type_String);
								bError = true;
								logger.error("key:" + key, ex);
							}
						} else if (thead.isLong()) {
							try {
								// 长整形时按字符串输出
								long d = ja.getLong(key);
								l = new Label(column++, row, d + "",
										BaseUtil.detFormat);
								sheet.addCell(l);

							} catch (Exception ex) {
								thead.setHeadType(IHeader.Type_String);
								bError = true;
								logger.error("key:" + key, ex);
							}

						} else {
							String str = ja.getString(key);
							if (doubleToStr.contains(thead.getDataIndex())) {
								// 数据库msisdn和imsi用的number型时为科学计数法，转换为正确的字符串
								if (!"".equals(str)) {
									java.text.NumberFormat nf = java.text.NumberFormat
											.getInstance();
									nf.setGroupingUsed(false);
									str = nf.format(Double.valueOf(str));
								}
							}

							if (str == null) {
								str = "";
							}
							l = new Label(column++, row, str,
									BaseUtil.detFormat);
							sheet.addCell(l);
						}

						if (bError) {
							String str = ja.getString(key);
							if (str == null) {
								str = "";
							}
							l = new Label(column++, row, str,
									BaseUtil.detFormat);
							sheet.addCell(l);
						}
					}
				}
			}

			int colIndex = 0;
			for (int i = 0; i < header.length; i++) {
				if (!header[i].isExport()) {
					continue;
				}
				if (header[i].isSetWidth()) {
					sheet.setColumnView(colIndex,
							header[i].getHeaderWidth() / 5);
				} else {
					sheet.setColumnView(colIndex, header[i].getHeaderName()
							.length() * 2);
				}
				colIndex++;
			}

			workbook.write();
			workbook.close();

		} catch (Exception e) {
			logger.error(e);
		}
	}
	
	/**
	 * 
	 * jxl.write.NumberFormat nf = new jxl.write.NumberFormat("#.##");
	 * jxl.write.WritableCellFormat wcfN = new jxl.write.WritableCellFormat(nf);
	 * jxl.write.Number labelNF = new jxl.write.Number(0, 2, 3.1415926, wcfN);
	 * sheet.addCell(labelNF);
	 * 
	 * Description: excel文件导出 //未使用
	 * @author lenovo
	 * @param out 文件输出流
	 * @param sheetName 工作表名
	 * @param header 表头
	 * @param dataIndex 与查询结果JSONObject的key对应的名称
	 * @param array 查询结果信息
	 * @param hiddens 列是否显示
	 * @param groupHeaders 表格列所占大小合并信息
	 * @exception IOException,RowsExceededException,WriteException
	 */
	public static void exportExcel(OutputStream out, String sheetName,
			String[] header, String[] dataIndex, JSONArray array,
			boolean[] hiddens,String[] groupHeaders) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportExcel be called !");
		}

		// 一些临时变量，用于写到excel中
		Label l = null;

		// add head
		int column = 0;
		int row = 0;

		WritableWorkbook workbook = null;

		try {

			workbook = Workbook.createWorkbook(out);

			WritableSheet sheet = workbook.createSheet(sheetName, 1);
			for (int i = 0; i < groupHeaders.length; i++) {
				
			}

			for (int i = 0; i < header.length; i++) {
				if (!hiddens[i]) {
					l = new Label(column++, row, header[i],
							BaseUtil.headerFormat);
					sheet.addCell(l);
				}
			}
			for (int i = 0; i < array.size(); i++) {
				column = 0;
				row = row + 1;

				if (i != 0 && i % PageSize == 0) {
					for (int is = 0; is < sheet.getColumns(); is++) {

						sheet.setColumnView(is, header.length);

					}
					int bh = i / PageSize;
					sheet = workbook.createSheet(String.valueOf(bh), bh);
					row = 0;
				}

				JSONObject ja = (JSONObject) array.get(i);
				if (ja == null) {
					continue;
				}
				int iw = 0;
				for (int index = 0; index < header.length; index++) {

					String key = dataIndex[index];
					String str = ja.getString(key);
					if (!hiddens[iw]) {

						if (str == null) {
							str = "";
						}
						l = new Label(column++, row, str, BaseUtil.detFormat);
						sheet.addCell(l);
					}
					iw = iw + 1;
				}
			}

			for (int i = 0; i < sheet.getColumns(); i++) {

				sheet.setColumnView(i, header.length);

			}

			workbook.write();
			workbook.close();

		} catch (Exception e) {
			logger.error(e);
		}
	}


	/**
	 * 
	 * Description: excel文件导出
	 * @author lenovo
	 * @param out 文件输出流
	 * @param sheetName 工作表名
	 * @param header 显示列的各种属性信息
	 * @param records 查询的结果信息
	 * @exception IOException,RowsExceededException,WriteException
	 */
	public static void exportExcel(OutputStream out, String sheetName,
			IHeader[] header, JSONArray records) {
		if (logger.isDebugEnabled()) {
			logger.debug("exportExcel be called !");
		}

		// 一些临时变量，用于写到excel中
		Label l = null;

		// add head
		int column = 0;
		int row = 0;

		WritableWorkbook workbook = null;

		try {

			workbook = Workbook.createWorkbook(out);
			WritableSheet sheet = workbook.createSheet(sheetName, 0);

			for (int i = 0; i < header.length; i++) {
				if (header[i].isExport()) {
					l = new Label(column++, row, header[i].getHeaderName(),
							BaseUtil.headerFormat);
					sheet.addCell(l);
				}
			}
			for (int i = 0; i < records.size(); i++) {
				column = 0;
				row = row + 1;
				if (i != 0 && i % PageSize == 0) {
					int colIndex = 0;
					for (int is = 0; is < header.length; is++) {
						if (!header[is].isExport()) {
							continue;
						}
						if (header[is].isSetWidth()) {
							sheet.setColumnView(colIndex,
									header[is].getHeaderWidth() / 5);
						} else {
							sheet.setColumnView(colIndex, header[is]
									.getHeaderName().length() * 2);
						}
						colIndex++;
					}
					int bh = i / PageSize;
					sheet = workbook.createSheet(String.valueOf(bh), bh);
					row = 0;
				}

				JSONObject ja = (JSONObject) records.get(i);
				if (ja == null) {
					continue;
				}
				int headIndex = 0;
				for (headIndex = 0; headIndex < header.length; headIndex++) {
					IHeader thead = header[headIndex];

					String key = thead.getDataIndex();
					boolean bError = false;
					if (thead.isExport()) {

						try {
							if ((thead.isDouble() || thead.isInt())
									&& ja.getDouble(key) > Integer.MAX_VALUE - 1) {
								thead.setHeadType(IHeader.Type_String);
							}
						} catch (Exception e) {
							// TODO Auto-generated catch block
							logger.error("The value is not a Double or int Value : "
									+ ja.getString(key) + " " + e);
						}

						if (thead.isDouble()) {
							try {
								WritableCellFormat df = getDoubleFormat();
								double d = ja.getDouble(key);
								Number nb = new jxl.write.Number(column, row,
										d, df);
								sheet.addCell(nb);
								column++;
							} catch (Exception ex) {
								thead.setHeadType(IHeader.Type_String);
								bError = true;
								logger.error(ex);
							}

						} else if (thead.isInt()) {
							try {
								WritableCellFormat df = getIntFormat();
								int d = ja.getInt(key);
								Number nb = new jxl.write.Number(column, row,
										d, df);
								sheet.addCell(nb);
								column++;
							} catch (Exception ex) {
								thead.setHeadType(IHeader.Type_String);
								bError = true;
								logger.error("key:" + key, ex);
							}
						} else if (thead.isLong()) {
							try {
								// 长整形时按字符串输出
								long d = ja.getLong(key);
								l = new Label(column++, row, d + "",
										BaseUtil.detFormat);
								sheet.addCell(l);

							} catch (Exception ex) {
								thead.setHeadType(IHeader.Type_String);
								bError = true;
								logger.error("key:" + key, ex);
							}

						} else {
							String str = ja.getString(key);
							if ("msisdn".equalsIgnoreCase(thead.getDataIndex())
									|| "imsi".equalsIgnoreCase(thead
											.getDataIndex())) {
								// 数据库msisdn和imsi用的number型时为科学计数法，转换为正确的字符串
								if (str.indexOf('E') != -1) {
									java.text.NumberFormat nf = java.text.NumberFormat
											.getInstance();
									nf.setGroupingUsed(false);
									str = nf.format(Double.valueOf(str));
								}
							}

							if (str == null) {
								str = "";
							}
							l = new Label(column++, row, str,
									BaseUtil.detFormat);
							sheet.addCell(l);
						}

						if (bError) {
							String str = ja.getString(key);
							if (str == null) {
								str = "";
							}
							l = new Label(column++, row, str,
									BaseUtil.detFormat);
							sheet.addCell(l);
						}
					}
				}
			}

			int colIndex = 0;
			for (int i = 0; i < header.length; i++) {
				if (!header[i].isExport()) {
					continue;
				}
				if (header[i].isSetWidth()) {
					sheet.setColumnView(colIndex,
							header[i].getHeaderWidth() / 5);
				} else {
					sheet.setColumnView(colIndex, header[i].getHeaderName()
							.length() * 2);
				}
				colIndex++;
			}

			workbook.write();
			workbook.close();

		} catch (Exception e) {
			logger.error(e);
		}
	}

	/**
	 * 
	 * Description: Txt文件导出
	 * @author lenovo
	 * @param writer
	 * @param fileName
	 * @param header 表头
	 * @param dataIndex 与查询结果JSONObject的key对应的名称
	 * @param array 查询的结果信息
	 * @param hiddens 列是否显示
	 * @exception IOException
	 */
	public static void exportTxt(Writer writer, String fileName,
			String[] header, String[] dataIndex, JSONArray array,
			boolean[] hiddens) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportTxt be called !");
		}

		if (header == null) {
			return;
		}
		StringBuffer sb = new StringBuffer();
		boolean first = true;
		for (int i = 0; i < header.length; i++) {
			if (!hiddens[i]) {
				if (first) {
					first = false;
				} else {

					sb.append(",");
				}
				sb.append(header[i]);
			}
		}

		sb.append("\r\n");

		for (int i = 0; i < array.size(); i++) {

			JSONObject ja = (JSONObject) array.get(i);
			if (ja == null) {
				continue;
			}
			first = true;

			int iw = 0;
			for (int index = 0; index < dataIndex.length; index++) {
				String key = dataIndex[index];
				String str = ja.getString(key);
				if (!hiddens[iw]) {

					if (str == null) {
						str = "";
					}
					if (first) {
						first = false;
					} else {
						sb.append(",");
					}
					sb.append(str);
				}
				iw = iw + 1;
			}

			sb.append("\r\n");
		}

		String re = sb.toString();

		try {
			writer.write(re);
		} catch (IOException e) {
			logger.error(e);
		}
	}

	/**
	 * 
	 * Description: Xml文件导出
	 * @author lenovo
	 * @param writer
	 * @param fileName
	 * @param header 表头
	 * @param dataIndex 与查询结果JSONObject的key对应的名称
	 * @param array 查询的结果信息
	 * @param hiddens 列是否显示
	 * @exception IOException
	 */
	public static void exportXml(Writer writer, String fileName,
			String[] header, String[] dataIndex, JSONArray array,
			boolean[] hiddens) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportXML be called !");
		}

		if (header == null) {
			return;
		}

		Document doc = DocumentHelper.createDocument();
		doc.setXMLEncoding("GBK");

		Element root = null;
		root = doc.addElement("mspweb");

		for (int i = 0; i < array.size(); i++) {

			JSONObject ja = (JSONObject) array.get(i);
			if (ja == null) {
				continue;
			}
			Element data = root.addElement("row");
			Iterator iter = ja.keys();
			int j = 0;
			int iw = 0;
			for (int index = 0; index < header.length; index++) {
				String key = dataIndex[index];
				String str = ja.getString(key);
				if (!hiddens[iw]) {
					if (str == null) {
						str = "";
					}
					data.addElement("td").addAttribute("name", header[j])
							.setText(str);
				}
				j++;
				iw = iw + 1;
			}

		}

		try {
			writer.write(doc.asXML());
		} catch (IOException e) {
			logger.error(e);
		}
	}

	/**
	 * 
	 * Description: Html文件导出
	 * @author lenovo
	 * @param writer
	 * @param fileName Htm页面table表格的标题
	 * @param header 表头
	 * @param dataIndex 与查询结果JSONObject的key对应的名称
	 * @param array 查询的结果信息
	 * @param hiddens 列是否显示
	 * @exception IOException
	 */
	public static void exportHtml(Writer writer, String fileName,
			String[] header, String[] dataIndex, JSONArray array,
			boolean[] hiddens) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportHTML be called !");
		}

		if (header == null) {
			return;
		}

		if (fileName == null) {
			fileName = "";
		}

		StringBuffer sb = new StringBuffer();

		sb.append("<html>");
		sb.append("<head>");
		sb.append("<style>");

		sb.append("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=gbk\" />");
		sb.append(" .tr_head{");
		sb.append("font-size: 12px;");
		sb.append("font-family: \"宋体\";");
		sb.append("font: bold;");
		sb.append("color:#FFFFFF;");
		sb.append("background-color: #5A8ECE;");
		sb.append("padding-top: 5px;");
		sb.append("text-align: center;");
		sb.append("height: 30px;");

		sb.append("}");

		sb.append(" .tr_row1 {" + "font-size: 12px;" + "font-family: \"宋体\";"
				+ "background-color: #f4f9ed;" + "padding-top: 3px;"
				+ "text-align: center;" + "cursor:hand;height: 25px;" + "}");

		sb.append(" td {");
		sb.append("font-size: 12px;");
		sb.append("font-family: \"宋体\";");
		sb.append("}");

		sb.append(" .td_row1 {" + "font-size: 12px;" + "font-family: \"宋体\";"
				+ "background-color:#ffffff;" + "padding-top: 3px;"
				+ "text-align: center;" + "cursor:hand;height: 25px;" + "}");

		sb.append("</style>");
		sb.append("</head>");
		sb.append("<body>");

		sb.append("<table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"1\" bgcolor=\"#CEDBEF\" >");
		sb.append("<caption><span class=\"tr_head\">" + fileName
				+ "<span></caption>");
		sb.append("<thead>");

		sb.append("<tr>");
		for (int i = 0; i < header.length; i++) {
			if (!hiddens[i]) {
				sb.append("<th  nowrap class=\"tr_head\">" + header[i]
						+ "</th>");
			}
		}
		sb.append("</tr>");

		for (int i = 0; i < array.size(); i++) {

			JSONObject ja = (JSONObject) array.get(i);
			if (ja == null) {
				continue;
			}

			int iw = 0;
			sb.append("<tr class=\"tr_row1\">");
			for (int index = 0; index < header.length; index++) {
				String key = dataIndex[index];
				String str = ja.getString(key);
				if (!hiddens[iw]) {
					if (str == null) {
						str = "";
					}
					sb.append("<td nowrap class=\"td_row1\">");
					sb.append(str);
					sb.append("</td>");
				}
				iw = iw + 1;
			}

			sb.append("</tr>");
		}

		sb.append("</thead>");
		sb.append("<tbody>");

		sb.append("</tbody>");
		sb.append("</table>");
		sb.append("</body>");
		sb.append("</html>");

		try {
			writer.write(sb.toString());
		} catch (IOException e) {
			logger.error(e);
		}
	}

	/**
	 * 
	 * Description: Pdf文件导出
	 * @author lenovo
	 * @param out 文件输出流
	 * @param fileName 工作表名
	 * @param header 表头
	 * @param dataIndex 与查询结果JSONObject的key对应的名称
	 * @param array 查询的结果信息
	 * @param hiddens 列是否显示
	 * @exception IOException,DocumentException
	 */
	public static void exportPdf(OutputStream out, String fileName,
			String[] header, String[] dataIndex, JSONArray array,
			boolean[] hiddens) {

		PdfGenerator
				.createPdf(out, fileName, header, dataIndex, array, hiddens);
	}


	/**
	 * 
	 * Description: Excel文件导出
	 * @author lenovo
	 * @param out 文件输出流
	 * @param sheetName 工作表名
	 * @param fieldList 表头显示字段的详细信息的list
	 * @param list 字段类型转换信息list
	 * @exception IOException,RowsExceededException,WriteException
	 */
	public static void exportExcel(FileOutputStream out, String sheetName,
			List<FieldInfo> fieldList, List<List<IData>> list) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportExcel be called ! sheetName:" + sheetName);
		}

		// 一些临时变量，用于写到excel中
		Label l = null;

		// add head
		int column = 0;
		int row = 0;

		WritableWorkbook workbook = null;

		try {

			workbook = Workbook.createWorkbook(out);
			WritableSheet sheet = workbook.createSheet(sheetName, 0);

			for (int i = 0; fieldList != null && i < fieldList.size(); i++) {

				l = new Label(column++, row, fieldList.get(i).getDesc(),
						BaseUtil.headerFormat);
				sheet.addCell(l);
			}

			String str = null;

			for (int i = 0; list != null && i < list.size(); i++) {
				column = 0;
				row = i + 1;

				List<IData> temp = (List<IData>) list.get(i);

				for (int j = 0; j < temp.size(); j++) {

					IData data = temp.get(j);
					if (data.getPrimitiveValue() == null) {
						str = "";
					} else {
						str = data.toStr();
					}
					l = new Label(column++, row, str, BaseUtil.detFormat);
					sheet.addCell(l);
				}

				if (i >= 65534) {
					break;
				}
			}

			for (int i = 0; i < sheet.getColumns(); i++) {

				sheet.setColumnView(i, fieldList.size());

			}

			workbook.write();
			workbook.close();

		} catch (Exception e) {
			logger.error(e);
		}
	}

	/**
	 * 
	 * Description: Txt文件导出
	 * @author lenovo
	 * @param writer
	 * @param sheetName  
	 * @param fieldList 表头显示字段的详细信息的list
	 * @param list 字段类型转换信息list
	 * @exception IOException
	 */
	public static void exportTxt(Writer writer, String sheetName,
			List<FieldInfo> fieldList, List<List<IData>> list) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportTxt be called !");
		}

		StringBuffer sb = new StringBuffer();
		boolean first = true;
		for (int i = 0; fieldList != null && i < fieldList.size(); i++) {

			if (first) {
				first = false;
			} else {

				sb.append(",");
			}
			sb.append(fieldList.get(i).getDesc());

		}

		sb.append("\r\n");

		String str = null;

		int num = 0;
		for (int i = 0; list != null && i < list.size(); i++) {

			List<IData> temp = list.get(i);

			for (int j = 0; j < temp.size(); j++) {

				first = true;

				if (first) {
					first = false;
				} else {
					sb.append(",");
				}
				IData data = temp.get(j);
				if (data.getPrimitiveValue() == null) {
					str = "";
				} else {
					str = data.toStr();
				}
				sb.append(str);
			}

			sb.append("\r\n");
			num++;
			if (num >= 100) {

				try {
					writer.write(sb.toString());
				} catch (Exception ex) {
					logger.error(ex);
				}
				sb = new StringBuffer();
			}
		}

		String re = sb.toString();

		try {
			writer.write(re);
			writer.close();
		} catch (IOException e) {
			logger.error(e);
		}
	}

	/**
	 * 
	 * Description: Xml文件导出
	 * @author lenovo 
	 * @param writer
	 * @param sheetName
	 * @param fieldList 表头显示字段的详细信息的list
	 * @param list 字段类型转换信息list
	 * @exception IOException
	 */
	public static void exportXml(Writer writer, String sheetName,
			List<FieldInfo> fieldList, List<List<IData>> list) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportXML be called !");
		}

		Document doc = DocumentHelper.createDocument();
		doc.setXMLEncoding("GBK");

		Element root = null;
		root = doc.addElement("mspweb");

		String str = null;
		for (int i = 0; list != null && i < list.size(); i++) {

			Element data = root.addElement("row");

			List<IData> temp = list.get(i);
			for (int j = 0; j < temp.size(); j++) {

				if (temp.get(j).getPrimitiveValue() == null) {
					str = "";
				} else {
					str = temp.get(j).toStr();
				}

				data.addElement("td")
						.addAttribute("name", fieldList.get(j).getDesc())
						.setText(str);
			}
		}

		try {
			writer.write(doc.asXML());
			writer.close();

		} catch (IOException e) {
			logger.error(e);
		}
	}

	/**
	 * 
	 * Description: Html文件导出
	 * @author lenovo
	 * @param writer
	 * @param sheetName Htm页面table表格的标题
	 * @param fieldList 表头显示字段的详细信息的list
	 * @param list 字段类型转换信息list
	 * @exception IOException
	 */
	public static void exportHtml(Writer writer, String sheetName,
			List<FieldInfo> fieldList, List<List<IData>> list) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportHTML be called !");
		}

		StringBuffer sb = new StringBuffer();

		sb.append("<html>");
		sb.append("<head>");
		sb.append("<style>");

		sb.append("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=gbk\" />");
		sb.append(" .tr_head{");
		sb.append("font-size: 12px;");
		sb.append("font-family: \"宋体\";");
		sb.append("font: bold;");
		sb.append("color:#FFFFFF;");
		sb.append("background-color: #5A8ECE;");
		sb.append("padding-top: 5px;");
		sb.append("text-align: center;");
		sb.append("height: 30px;");

		sb.append("}");

		sb.append(" .tr_row1 {" + "font-size: 12px;" + "font-family: \"宋体\";"
				+ "background-color: #f4f9ed;" + "padding-top: 3px;"
				+ "text-align: center;" + "cursor:hand;height: 25px;" + "}");

		sb.append(" td {");
		sb.append("font-size: 12px;");
		sb.append("font-family: \"宋体\";");
		sb.append("}");

		sb.append(" .td_row1 {" + "font-size: 12px;" + "font-family: \"宋体\";"
				+ "background-color:#ffffff;" + "padding-top: 3px;"
				+ "text-align: center;" + "cursor:hand;height: 25px;" + "}");

		sb.append("</style>");
		sb.append("</head>");
		sb.append("<body>");

		sb.append("<table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"1\" bgcolor=\"#CEDBEF\" >");
		sb.append("<caption><span class=\"tr_head\">" + sheetName
				+ "<span></caption>");
		sb.append("<thead>");

		sb.append("<tr>");
		for (int i = 0; i < fieldList.size(); i++) {
			sb.append("<th  nowrap class=\"tr_head\">"
					+ fieldList.get(i).getDesc() + "</th>");
		}
		sb.append("</tr>");

		String str = null;
		int num = 0;
		for (int i = 0; list != null && i < list.size(); i++) {

			int j = 0;
			sb.append("<tr class=\"tr_row1\">");

			List<IData> temp = list.get(i);

			for (j = 0; j < temp.size(); j++) {

				IData data = temp.get(j);
				if (data.getPrimitiveValue() == null) {
					str = "";
				} else {
					str = data.toStr();
				}
				sb.append("<td nowrap class=\"td_row1\">");
				sb.append(str);
				sb.append("</td>");
			}
			sb.append("</tr>");
			num++;
			if (num >= 100) {

				try {
					writer.write(sb.toString());
				} catch (IOException e) {
					logger.error(e);
				}
				sb = new StringBuffer();
			}
		}

		sb.append("</thead>");
		sb.append("<tbody>");

		sb.append("</tbody>");
		sb.append("</table>");
		sb.append("</body>");
		sb.append("</html>");

		try {
			writer.write(sb.toString());
			writer.close();

		} catch (IOException e) {
			logger.error(e);
		}
	}

	/**
	 * 
	 * Description: 导出模块管理的insert语句
	 * @author lenovo 
	 * @param writer
	 * @param fileName
	 * @param header 表头
	 * @param dataIndex 与查询结果JSONObject的key对应的名称
	 * @param array 查询的结果信息
	 * @param hiddens 列是否显示
	 * @exception IOException
	 */
	public static void exportTxt3(Writer writer, String fileName,
			String[] header, String[] dataIndex, JSONArray array,
			boolean[] hiddens) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportTxt be called !");
		}

		if (header == null) {
			return;
		}
		StringBuffer sb = new StringBuffer();

		for (int i = 0; i < array.size(); i++) {
			boolean first = true;
			sb.append("insert into syu_module_list (");

			for (int j = 1; j < dataIndex.length; j++) {
				if (!hiddens[j]) {
					if (first) {
						first = false;
					} else {

						sb.append(",");
					}
					sb.append(dataIndex[j]);
				}
			}

			sb.append(") values (");

			JSONObject ja = (JSONObject) array.get(i);
			if (ja == null) {
				continue;
			}
			first = true;

			int iw = 1;
			for (int index = 1; index < dataIndex.length; index++) {
				String key = dataIndex[index];
				String str = ja.getString(key);
				if (!hiddens[iw]) {

					if (first) {
						first = false;
					} else {
						sb.append(",");
					}
					sb.append("'");
					sb.append(str);
					sb.append("'");
				}
				iw = iw + 1;
			}

			sb.append(");\r\n");
		}

		String re = sb.toString();

		try {
			writer.write(re);
		} catch (IOException e) {
			logger.error(e);
		}
	}

	/**
	 * 
	 * Description: Pdf文件导出
	 * @author lenovo 
	 * @param out 文件输出流
	 * @param sheetName Pdf文件标题
	 * @param fieldList 表头显示字段的详细信息的list
	 * @param list 字段类型转换信息list
	 * @exception IOException,DocumentException
	 */
	public static void exportPdf(OutputStream out, String sheetName,
			List<FieldInfo> fieldList, List<List<IData>> list) {

		PdfGenerator.createPdf(out, sheetName, fieldList, list);
	}

	/**
	 * 
	 * Description: Imei文件导出
	 * @author lenovo 
	 * @param writer
	 * @param fileName
	 * @param header 表头
	 * @param array 查询的结果信息
	 * @param hiddens 列是否显示
	 * @exception IOException
	 */
	public static void exportImei(Writer writer, String fileName,
			String[] header, JSONArray array, boolean[] hiddens) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportTxt be called !");
		}

		if (header == null) {
			return;
		}
		StringBuffer sb = new StringBuffer();
		boolean first = true;
		for (int i = 1; i < header.length; i++) {
			if (!hiddens[i]) {
				if (first) {
					first = false;
				} else {

					sb.append(",");
				}
				sb.append(header[i]);
			}
		}

		sb.append("\r\n");

		for (int i = 0; i < array.size(); i++) {

			JSONObject ja = (JSONObject) array.get(i);
			if (ja == null) {
				continue;
			}
			first = true;

			Iterator iter = ja.keys();
			int iw = 0;
			while (iter.hasNext()) {
				String key = (String) iter.next();
				String str = ja.getString(key);
				if (!hiddens[iw]) {

					if ("id".equals(key))
						continue;

					if (str == null) {
						str = "";
					}
					if (first) {
						first = false;
					} else {
						sb.append(",");
					}
					sb.append(str);
				}
				iw = iw + 1;
			}

			sb.append("\r\n");
		}

		String re = sb.toString();

		try {
			writer.write(re);
		} catch (IOException e) {
			logger.error(e);
		}
	}

	/**
	 * 
	 * Description: ImeiSql文件导出
	 * @author lenovo 
	 * @param writer
	 * @param fileName
	 * @param header 
	 * @param array 查询的结果信息
	 * @param hiddens
	 * @exception IOException
	 */
	public static void exportImeiSql(Writer writer, String fileName,
			String[] header, JSONArray array, boolean[] hiddens) {

		if (logger.isDebugEnabled()) {
			logger.debug("exportTxt be called !");
		}

		if (header == null) {
			return;
		}
		StringBuffer sb = new StringBuffer();
		boolean first = true;

		for (int i = 0; i < array.size(); i++) {

			JSONObject ja = (JSONObject) array.get(i);
			if (ja == null) {
				continue;
			}
			first = true;

			Iterator iter = ja.keys();
			int iw = 0;
			while (iter.hasNext()) {
				String key = (String) iter.next();
				String str = ja.getString(key);
				if (!hiddens[iw]) {

					if ("id".equals(key))
						continue;

					if (str == null) {
						str = "";
					}
					if (first) {
						first = false;
						sb.append("INSERT INTO sys_tac_config (tac,mobile_vendor,mobile_type,mobile_class,mobile_os,network_mode) VALUES ('");
					} else {
						if ("network_mode".equals(key) || "mobile_class".equals(key))
							sb.append("',");
						else if ("mobile_os".equals(key))
							sb.append(",");
						else
							sb.append("','");
					}
					sb.append(str);
				}
				iw = iw + 1;
			}

			sb.append(");\r\n");
		}


		String re = sb.toString();

		try {
			writer.write(re);
		} catch (IOException e) {
			logger.error(e);
		}
	}

	/**
	 * 
	 * Description: Imei属性文件导出
	 * @author lenovo 
	 * @param writer
	 * @param fileName
	 * @param header
	 * @param array 查询的结果信息
	 * @param hiddens
	 * @exception IOException
	 */
	public static void exportImeiProperties(Writer writer, String fileName,
			String[] header, JSONArray array, boolean[] hiddens) {
		if (logger.isDebugEnabled()) {
			logger.debug("exportTxt be called !");
		}

		if (header == null) {
			return;
		}
		StringBuffer sb = new StringBuffer();
		boolean first = true;
		for (int i = 1; i < header.length; i++) {
			if (!hiddens[i]) {
				if (first) {
					sb.append("#");
					first = false;
				} else {

					sb.append(",");
				}
				sb.append(header[i]);
			}
		}

		sb.append("\r\n");

		for (int i = 0; i < array.size(); i++) {

			JSONObject ja = (JSONObject) array.get(i);
			if (ja == null) {
				continue;
			}
			first = true;

			Iterator iter = ja.keys();
			int iw = 0;
			while (iter.hasNext()) {
				String key = (String) iter.next();
				String str = ja.getString(key);
				if (!hiddens[iw]) {

					if ("id".equals(key))
						continue;

					if (str == null) {
						str = "";
					}
					if (first) {
						first = false;
					} else {
						sb.append(",");
					}
					sb.append(str.replace(".0", ""));
				}
				iw = iw + 1;
			}

			sb.append("\r\n");
		}

		String re = sb.toString();

		try {
			writer.write(re);
		} catch (IOException e) {
			logger.error(e);
		}
	}
}
