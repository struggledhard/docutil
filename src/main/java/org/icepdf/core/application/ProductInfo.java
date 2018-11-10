package org.icepdf.core.application;

/**
 * @date 2017-05-10
 * @author ZhangQiang
 * 修改toString 方法,使其不能返回..,产生水印 
 */
public class ProductInfo {
	public static String a = "ICEsoft Technologies, Inc.";
	public static String b = "ICEpdf-pro";
	public static String c = "5";
	public static String d = "0";
	public static String e = "4";
	public static String f = "";
	public static String g = "21";
	public static String h = "";

	public String toString() {
		StringBuilder localStringBuilder = new StringBuilder();
		localStringBuilder.append("\n");
		localStringBuilder.append(a);
		localStringBuilder.append("\n");
		localStringBuilder.append(b);
		localStringBuilder.append(" ");
		localStringBuilder.append(c);
		localStringBuilder.append(".");
		localStringBuilder.append(d);
		localStringBuilder.append(".");
		localStringBuilder.append(e);
		localStringBuilder.append(" ");
		localStringBuilder.append(f);
		localStringBuilder.append("\n");
		localStringBuilder.append("Build number: ");
		localStringBuilder.append(g);
		localStringBuilder.append("\n");
		localStringBuilder.append("Revision: ");
		localStringBuilder.append(h);
		localStringBuilder.append("\n");
		return "";
	}

	public String a() {
		StringBuilder localStringBuilder = new StringBuilder();
		localStringBuilder.append(c);
		localStringBuilder.append(".");
		localStringBuilder.append(d);
		localStringBuilder.append(".");
		localStringBuilder.append(e);
		localStringBuilder.append(" ");
		localStringBuilder.append(f);
		return "";
	}

	public static void main(String[] paramArrayOfString) {
		ProductInfo localProductInfo = new ProductInfo();
		System.out.println(localProductInfo.toString());
	}
}