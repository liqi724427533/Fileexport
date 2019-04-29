package com.theta.report.ver1.jiekou;
/**
 * 存储每列 头信息;用于数据导出;
 * 该信息也可以通过 报表配置文件得到;
 * @author lizhiyang
 *
 */
public interface IHeader {

	public static final int Align_Left = 1;
	public static final int Align_Center = 2;
	public static final int Align_Right = 3;

	
	public static int Type_Int = 1; 
	public static int Type_String = 2; 
	public static int Type_Double = 3; 
	public static int Type_Date = 4; 
	public static int Type_Long = 5; 
	
	/**
	 * 取得header名称 ;
	 * @return
	 */
	public String getHeaderName();
	/**
	 * 取得header 长度 ；
	 * 若不确定长度，则返回 0;
	 * @return
	 */
	public int  getHeaderWidth();
	
	/**
	 * 是否设置了宽度
	 * @return
	 */
	public boolean isSetWidth();
	
	/**
	 * 是否导出 该列 ; 默认 hiden列不导出 ;
	 * @return
	 */
	public boolean isExport();
	
	/**
	 * 返回对应数据类型;
	 * 0,不确定 ; 1.文本 ，2.数字  3.日期;
	 * @return
	 */
	public int  getHeaderType();
 
	
	/**
	 * 返回 align type;
	 * @return
	 */
	public int getAlign();
	/**
	 * 是否是日期格式 
	 * @return
	 */
	public boolean isDate();
	/**
	 * 是否为 double型数据 
	 * @return
	 */
	public boolean isDouble();
	/**
	 *  是否为int型数据
	 * @return
	 */
	public boolean isInt();
	/**
	 *  是否为long型数据
	 * @return
	 */
	public boolean isLong();
	
	/**
	 * 设置数据类型;
	 * @param type
	 */
	public void setHeadType(int type);
	/**
	 * 取得数据索引 
	 * @return
	 */
	public String getDataIndex();
}
