package util.tingfeng.java.files.officeToPdf;

import java.io.File;

import util.tingfeng.java.files.officeToPdf.OfficeToPdfFactory.OfficeType;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComFailException;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
/**
 * 适用于WPS2015
 * @author dview76
 *
 */
public class OfficeToPdfByWps implements OfficeToPdfI{
	public final static String WORDSERVER_STRING="KWPS.Application";
	public final static String PPTSERVER_STRING="KWPP.Application";
	public final static String EXECLSERVER_STRING="KET.Application";
	private static final int wdFormatPDF = 17;
	private static final int xlTypePDF = 0;
	private static final int ppSaveAsPDF = 32;
	
	/**
	 * @return 操作成功与否的提示信息. 如果返回 -1, 表示找不到源文件, 或url.properties配置错误; 如果返回 0,
	 *         则表示操作成功; 返回1, 则表示转换失败
	 */
	@Override
	public int officeToPdf(OfficeToPDFInfo officeToPDFInfo) {
		String sourceFile=officeToPDFInfo.sourceUrl;
		String destFile=officeToPDFInfo.destUrl;
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
			if(extentionName.equalsIgnoreCase("ppt")||extentionName.equalsIgnoreCase("pptx")||extentionName.equalsIgnoreCase("wpt"))
			{
				ppt2pdf(officeToPDFInfo.sourceUrl,officeToPDFInfo.destUrl);
			}else if(extentionName.equalsIgnoreCase("doc")||extentionName.equalsIgnoreCase("docx")||extentionName.equalsIgnoreCase("wps")){
				doc2pdf(officeToPDFInfo.sourceUrl,officeToPDFInfo.destUrl);
			}else if(extentionName.equalsIgnoreCase("xls")||extentionName.equalsIgnoreCase("xlsx")||extentionName.equalsIgnoreCase("et")){
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
		 ActiveXComponent pptActiveXComponent=null; 
		 ActiveXComponent workbook = null; 
		try {
        	 ComThread.InitSTA();//初始化COM线程  
             pptActiveXComponent = new ActiveXComponent(WORDSERVER_STRING);//初始化exe程序  
             Variant[] openParams=new Variant[]{
            		new Variant(srcFilePath),//filePath
            		new Variant(true),
            		new Variant(true)//readOnley
             };
             workbook = pptActiveXComponent.invokeGetComponent("Documents").invokeGetComponent
            		("Open",openParams);
             workbook.invoke("SaveAs",new Variant(pdfFilePath),new Variant(wdFormatPDF));                
             return true;
		}finally{
        	 if(workbook!=null)
        	 {
        		 workbook.invoke("Close"); 
        		 workbook.safeRelease(); 
        	 }
        	 if(pptActiveXComponent!=null)
        	 {  		
        		 pptActiveXComponent.invoke("Quit"); 
        		 pptActiveXComponent.safeRelease();
        	 }
        	 ComThread.Release();  
         }
    }
	
	protected static boolean ppt2pdf(String srcFilePath, String pdfFilePath) {
		 ActiveXComponent pptActiveXComponent=null; 
		 ActiveXComponent workbook = null;  
	     boolean readonly = true;
		try {
        	 ComThread.InitSTA();//初始化COM线程  
             pptActiveXComponent = new ActiveXComponent(PPTSERVER_STRING);//初始化exe程序  
             workbook = pptActiveXComponent.invokeGetComponent("Presentations").invokeGetComponent
            		("Open",new Variant(srcFilePath),new Variant(readonly));
             workbook.invoke("SaveAs",new Variant(pdfFilePath),new Variant(ppSaveAsPDF));                
             return true;
		}finally{
        	 if(workbook!=null)
        	 {
        		 workbook.invoke("Close"); 
        		 workbook.safeRelease(); 
        	 }
        	 if(pptActiveXComponent!=null)
        	 {  		
        		 pptActiveXComponent.invoke("Quit"); 
        		 pptActiveXComponent.safeRelease();
        	 }
        	 ComThread.Release();  
         }
    } 
	 public static boolean excel2PDF(String srcFilePath,String pdfFilePath){
		ActiveXComponent et = null; 
	    Dispatch workbooks = null;  
	    Dispatch workbook = null;  
	         ComThread.InitSTA();//初始化COM线程  
	         //ComThread.InitSTA(true);  
	         try {  
	             et = new ActiveXComponent(EXECLSERVER_STRING);//初始化et.exe程序  
	             et.setProperty("Visible", new Variant(false));  
	             workbooks = et.getProperty("Workbooks").toDispatch();  
	             //workbook = Dispatch.call(workbooks, "Open", filename).toDispatch();//这一句也可以的  
	             workbook = Dispatch.invoke(workbooks,"Open",Dispatch.Method,new Object[]{srcFilePath,0,true},new int[1]).toDispatch();   
	             //Dispatch.invoke(workbook,"SaveAs",Dispatch.Method,new Object[]{pdfFilePath,xlTypePDF},new int[1]);
	             Dispatch.call(workbook,"ExportAsFixedFormat",new Object[]{xlTypePDF,pdfFilePath});
	             return true;
	         }finally{
	        	 if(workbook!=null)
	        	 {
	        		 Dispatch.call(workbook,"Close");
	        		 workbook.safeRelease(); 
	        	 }
	        	 if(et!=null)
	        	 {  		
	        		 et.invoke("Quit"); 
	        		 et.safeRelease();
	        	 }
	        	 ComThread.Release();  
	         }
	 }
	@Override
	public OfficeType getOfficeType() {
		return OfficeType.Wps;
	}

}
