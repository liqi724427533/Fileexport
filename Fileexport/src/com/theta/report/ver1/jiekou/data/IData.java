package com.theta.report.ver1.jiekou.data;

import java.math.BigInteger;

/**
 * 包装 数据库查询出的数据。
 * @author Administrator
 *
 */
public interface IData {

	/**
	 * 合并数据 ; 数值型数据相加；字符串不相加;
	 * @param obj
	 */
	public void add(IData obj);
	
	/**
	 * this/obj;
	 * @param obj
	 */
	public double divide(Object obj);
	
	/**
	 * 连接对象 
	 * @param obj
	 */
	public void concat(IData obj);
	
	/**
	 * 取得 原生对象 , Integer, Double, String , Long ,.....
	 * @return
	 */
	public Object getPrimitiveValue();
	
	/**
	 * 返回对象字符串表示形式 ;
	 * null返回 ""; 字符串返回 "null"
	 * @return
	 */
	public String toStr();

	/**
	 * 取得格式化后的原始对象，double 保留两位小数
	 * @return
	 */
	public Object getFormatPrimitiveValue();

	/**
	 * 返回长整形数
	 * @return
	 */
	public Long getLongPrimitiveValue();

	/**
	 * 排序算法。
	 * @param data2
	 * @return
	 */
	public int comp(IData data2);

	/**
	 * 取得对应big integer 值
	 * @return
	 */
	public BigInteger getBigValue();
	
}
