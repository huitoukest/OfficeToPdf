package util.tingfeng.java.files.officeToPdf;
public class FileUtils {
		/**
		 * 
		 * @param filePath
		 * @return 返回文件的扩展名,如果扩展名不存在返回"",否则返回原值;
		 * 返回的扩展名不包含小点；
		 */
	    public static String getFileExtension(String filePath){
	    	if(filePath==null||filePath.trim().length()<1)
	    		return filePath;
	    	filePath=filePath.toLowerCase();
	    	int dotIndex = filePath.lastIndexOf(".");
	    	if (dotIndex <=0||(dotIndex+1==filePath.length())) {
		            return "";
		        }else{
		        	return filePath.substring(dotIndex+1, filePath.length());
		        }
	    }   
}
