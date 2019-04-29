package com.theta.common.app.util;

import java.io.ByteArrayOutputStream;
import java.io.DataOutput;
import java.io.DataOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.log4j.Logger;

import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.theta.report.ver1.dim1.model.data.FieldInfo;
import com.theta.report.ver1.jiekou.data.IData;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

public class PdfGenerator {

	protected static Logger logger = Logger.getLogger(PdfGenerator.class);

	/**
	 * Description: 创建生成Pdf文件
	 * @param out 文件输出流
	 * @param fileName pdf文件标题
	 * @param header 表头
	 * @param dataIndex 与查询结果JSONObject的key对应的名称
	 * @param array 查询结果信息
	 * @param hiddens 列是否显示
	 * @return 文件大小
	 * @exception IOException,DocumentException
	 */
	public static int createPdf(OutputStream out, String fileName,
			String[] header, String[] dataIndex, JSONArray array,
			boolean[] hiddens) {

		if (header == null) {
			return 0;
		}
		if (fileName == null) {
			fileName = "";
		}

		Rectangle rectPageSize = new Rectangle(PageSize.A4);
		rectPageSize.rotate();
		Document document = new Document(rectPageSize, 25, 25, 50, 25);
		document.addTitle(fileName);
		document.addAuthor("theta");
		document.addSubject(fileName);

		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try {

			// BaseFont bfChinese = BaseFont.createFont("STSong-Light",
			// "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);

			String path = getpath();
			logger.debug(path);
			BaseFont bfChinese = BaseFont.createFont(path+"font/SIMYOU.TTF",
					BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

			Font fontChinese = new Font(bfChinese, 12, Font.NORMAL);
			PdfWriter.getInstance(document, os);
			document.open();
			int i_hiddens = 0;
			for (int i = 0; i < hiddens.length; i++) {
				if (!hiddens[i]) {
					i_hiddens = i_hiddens + 1;
				}
			}
			PdfPTable table = new PdfPTable(i_hiddens);

			for (int i = 0; i < header.length; i++) {
				if (!hiddens[i]) {
					PdfPCell cell = new PdfPCell();
					cell.addElement(new Paragraph(header[i], fontChinese));
					table.addCell(cell);
				}
			}

			for (int i = 0; i < array.size(); i++) {

				JSONObject ja = (JSONObject) array.get(i);
				if (ja == null) {
					continue;
				}

				int j = 0;
				int iw = 0;
				for (int index = 0; index < header.length; index++) {
					String key = dataIndex[index];
					String str = ja.getString(key);
					if (!hiddens[iw]) {
						if (str == null) {
							str = "";
						}
						if ("msisdn".equalsIgnoreCase(key)
								|| "imsi".equalsIgnoreCase(key)
								|| key.indexOf("total") != -1
								|| key.indexOf("_num") != -1
								|| key.indexOf("bytes") != -1) {
							// 数据库msisdn和imsi用的number型时为科学计数法，转换为正确的字符串
							if (str.indexOf('E') != -1) {
								java.text.NumberFormat nf = java.text.NumberFormat
										.getInstance();
								nf.setGroupingUsed(false);
								str = nf.format(Double.valueOf(str));
							}
						}
						PdfPCell cell = new PdfPCell();
						cell.addElement(new Paragraph(str, fontChinese));
						table.addCell(cell);

					}
					iw = iw + 1;
				}

			}

			document.add(table);

			// document.add(new Paragraph("世界你好！", fontChinese));
		} catch (Exception e) {
			logger.error(e);
		} finally {
			document.close();

		}

		DataOutput output = new DataOutputStream(out);

		byte[] bytes = os.toByteArray();
		for (int num = 0; num < bytes.length; num++) {
			try {
				output.write(bytes[num]);
			} catch (IOException e) {
				logger.error(e);
			}
		}
		return bytes.length;
	}

	/**
	 * 
	 * Description: 创建生成Pdf文件
	 * @param out 文件输出流
	 * @param sheetName pdf文件标题
	 * @param fieldList 表头显示字段的详细信息的list
	 * @param list 字段类型转换信息list
	 * @return 文件大小
	 * @exception IOException,DocumentException
	 */
	public static int createPdf(OutputStream out, String sheetName,
			List<FieldInfo> fieldList, List<List<IData>> list) {

		Rectangle rectPageSize = new Rectangle(PageSize.A4);
		rectPageSize.rotate();
		Document document = new Document(rectPageSize, 25, 25, 50, 25);
		document.addTitle(sheetName);
		document.addAuthor("theta");
		document.addSubject(sheetName);

		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try {

			// BaseFont bfChinese = BaseFont.createFont("STSong-Light",
			// "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);

			String path = getpath();
			logger.debug(path);
			BaseFont bfChinese = BaseFont.createFont("SIMYOU.TTF",
					BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

			Font fontChinese = new Font(bfChinese, 12, Font.NORMAL);
			PdfWriter.getInstance(document, os);
			document.open();

			PdfPTable table = new PdfPTable(fieldList.size());

			for (int i = 0; fieldList != null && i < fieldList.size(); i++) {
				PdfPCell cell = new PdfPCell();
				cell.addElement(new Paragraph(fieldList.get(i).getDesc(),
						fontChinese));
				table.addCell(cell);
			}

			String str = null;
			for (int i = 0; list != null && i < list.size(); i++) {

				List<IData> temp = list.get(i);

				int j = 0;

				for (j = 0; j < temp.size(); j++) {

					IData data = temp.get(j);

					if (data.getPrimitiveValue() == null) {
						str = "";
					} else {
						str = data.toStr();
					}
					PdfPCell cell = new PdfPCell();
					cell.addElement(new Paragraph(str, fontChinese));
					table.addCell(cell);
				}
			}

			document.add(table);

			// document.add(new Paragraph("世界你好！", fontChinese));
		} catch (Exception e) {
			logger.error(e);
		} finally {
			document.close();

		}

		DataOutput output = new DataOutputStream(out);

		byte[] bytes = os.toByteArray();
		for (int num = 0; num < bytes.length; num++) {
			try {
				output.write(bytes[num]);
			} catch (IOException e) {
				logger.error(e);
			}
		}
		return bytes.length;
	}

	
	public static String getpath() {
		String path = PdfGenerator.class.getClassLoader().getResource("")
				.getPath();
		path = path.substring(0, path.length() - 8);
		return path;
	}
}
