package util.tingfeng.java.files.officeToPdf;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.ConnectException;

import util.tingfeng.java.files.officeToPdf.OfficeToPdfFactory.OfficeType;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

public class OfficeToPdfByOpenOffice  implements OfficeToPdfI{

	@Override
	public int officeToPdf(OfficeToPDFInfo officeToPDFInfo) {
		String sourceFile=officeToPDFInfo.sourceUrl;
		String destFile=officeToPDFInfo.destUrl;
		String OpenOffice_HOME=officeToPDFInfo.openOfficeHOME;
		try {
			File inputFile = new File(sourceFile);
			if (!inputFile.exists()) {
				return -1;// 找不到源文件, 则返回-1
			}

			// 如果目标路径不存在, 则新建该路径
			File outputFile = new File(destFile);
			if (!outputFile.getParentFile().exists()) {
				outputFile.getParentFile().mkdirs();
			}

			//= "D:\\Program Files\\OpenOffice.org 3";//这里是OpenOffice的安装目录, 在我的项目中,为了便于拓展接口,没有直接写成这个样子,但是这样是绝对没问题的
			// 如果从文件中读取的URL地址最后一个字符不是 '\'，则添加'\'
			if (OpenOffice_HOME.charAt(OpenOffice_HOME.length() - 1) != '\\') {
				OpenOffice_HOME += "\\";
			}
			// 启动OpenOffice的服务
			String command = OpenOffice_HOME
					+ "program\\soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;StarOffice.ServiceManager\" -nofirststartwizard";
			Process pro = Runtime.getRuntime().exec(command);
			// connect to an OpenOffice.org instance running on port 8100
			OpenOfficeConnection connection = new SocketOpenOfficeConnection(
					"127.0.0.1", 8100);
			connection.connect();
			// convert
			DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
			//DocumentConverter converter = new StreamOpenOfficeDocumentConverter(connection); 
			
			converter.convert(inputFile, outputFile);
			// close the connection
			connection.disconnect();
			// 关闭OpenOffice服务的进程
			pro.destroy();
			return 0;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return -1;
		} catch (ConnectException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return 1;
	}
	@Override
	public OfficeType getOfficeType() {
		return OfficeType.OpenOffice;
	}
}
