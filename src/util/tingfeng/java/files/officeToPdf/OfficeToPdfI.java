package util.tingfeng.java.files.officeToPdf;

import util.tingfeng.java.files.officeToPdf.OfficeToPdfFactory.OfficeType;

public interface OfficeToPdfI {	
	/**
	 * 操作成功与否的提示信息. 如果返回 -1, 表示找不到源文件, 或url.properties配置错误; 如果返回 0,
	 * 则表示操作成功; 返回1, 则表示转换失败
	 * @param officeToPDFInfo
	 * @return
	 */
	public int officeToPdf(OfficeToPDFInfo officeToPDFInfo);
	
	public OfficeType getOfficeType();
}
