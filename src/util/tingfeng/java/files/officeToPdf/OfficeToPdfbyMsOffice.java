package util.tingfeng.java.files.officeToPdf;

import java.io.File;

import util.tingfeng.java.files.officeToPdf.OfficeToPdfFactory.OfficeType;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComFailException;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class OfficeToPdfbyMsOffice implements OfficeToPdfI{

	private static final int wdFormatPDF = 17;
	private static final int xlTypePDF = 0;
	private static final int ppSaveAsPDF = 32;
	//private static final int msoTrue = -1;
	//private static final int msofalse = 0;
	/**
	 * @return 操作成功与否的提示信息. 如果返回 -1, 表示找不到源文件, 或url.properties配置错误; 如果返回 0,
	 *         则表示操作成功; 返回1, 则表示转换失败
	 */
	@Override
	public int officeToPdf(OfficeToPDFInfo officeToPDFInfo) {
		String sourceFile=officeToPDFInfo.sourceUrl;
		String destFile=officeToPDFInfo.destUrl;
		//String OpenOffice_HOME=officeToPDFInfo.openOfficeHOME;
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
			String extentionName=FileUtils.getFileExtension(sourceFile);
			if(extentionName.equalsIgnoreCase("ppt")||extentionName.equalsIgnoreCase("pptx"))
			{
				ppt2pdf(officeToPDFInfo.sourceUrl,officeToPDFInfo.destUrl);
			}else if(extentionName.equalsIgnoreCase("doc")||extentionName.equalsIgnoreCase("docx")){
				doc2pdf(officeToPDFInfo.sourceUrl,officeToPDFInfo.destUrl);
			}else if(extentionName.equalsIgnoreCase("xls")||extentionName.equalsIgnoreCase("xlsx")){
				excel2PDF(officeToPDFInfo.sourceUrl,officeToPDFInfo.destUrl);
			}	
			return 0;
		}catch (ComFailException e) { 
			e.printStackTrace();
            return 1;  
        } catch(Exception e) {
			e.printStackTrace();
			return 1;
		}
	}
	
	protected static boolean doc2pdf(String srcFilePath, String pdfFilePath) {  
        ActiveXComponent app = null;  
        Dispatch doc = null;  
        try {  
            ComThread.InitSTA();  
            app = new ActiveXComponent("Word.Application");  
            app.setProperty("Visible", false);  
            Dispatch docs = app.getProperty("Documents").toDispatch();  
            doc = Dispatch.invoke(docs, "Open", Dispatch.Method,  
                    new Object[] { srcFilePath,   
                                                 new Variant(false),   
                                                 new Variant(true),//是否只读  
                                                 new Variant(false),   
                                                 new Variant("pwd") },  
                    new int[1]).toDispatch();  
            // Dispatch.put(doc, "Compatibility", false);  //兼容性检查,为特定值false不正确  
            Dispatch.put(doc, "RemovePersonalInformation", false);  
            Dispatch.call(doc, "ExportAsFixedFormat", pdfFilePath,wdFormatPDF); // word保存为pdf格式宏，值为17
            return true; // set flag true;  
        }finally {  
            if (doc != null) {  
                Dispatch.call(doc, "Close", false);  
            }  
            if (app != null) {  
                app.invoke("Quit", 0);  
            }  
            ComThread.Release();  
        }  
    }
	
	protected static boolean ppt2pdf(String srcFilePath, String pdfFilePath) {  
        ActiveXComponent app = null;  
        Dispatch ppt = null;  
            try {  
                ComThread.InitSTA();  
                app = new ActiveXComponent("PowerPoint.Application");  
                Dispatch ppts = app.getProperty("Presentations").toDispatch();  
  
                // 因POWER.EXE的发布规则为同步，所以设置为同步发布  
                ppt = Dispatch.call(ppts, "Open", srcFilePath, true,// ReadOnly  
                        true,// Untitled指定文件是否有标题  
                        false// WithWindow指定文件是否可见  
                        ).toDispatch();  
                Dispatch.call(ppt, "SaveAs", pdfFilePath, ppSaveAsPDF); //ppSaveAsPDF为特定值32    
                return true; // set flag true;  
            }finally {  
                if (ppt != null) {  
                    Dispatch.call(ppt, "Close");  
                }  
                if (app != null) {  
                    app.invoke("Quit");  
                }  
                ComThread.Release();  
            }
    } 
	 public static boolean excel2PDF(String inputFile,String pdfFile){
	     ActiveXComponent app=null;
	     Dispatch excel =null;
		 try{
	        ComThread.InitSTA();  
	        app = new ActiveXComponent("Excel.Application");
	        app.setProperty("Visible", false);
	        Dispatch excels = app.getProperty("Workbooks").toDispatch();
	        excel = Dispatch.call(excels,"Open",inputFile, false,true).toDispatch();
	        Dispatch.call(excel,"ExportAsFixedFormat", xlTypePDF,pdfFile);
	        return true;
	    }finally{
	    	  if (excel!= null) {
	    		  Dispatch.call(excel, "Close");  
              }  
              if (app != null) {  
                  app.invoke("Quit");  
              }
              ComThread.Release();  
	    }	         
	    }
	@Override
	public OfficeType getOfficeType() {
		return OfficeType.MsOffice;
	}

}
