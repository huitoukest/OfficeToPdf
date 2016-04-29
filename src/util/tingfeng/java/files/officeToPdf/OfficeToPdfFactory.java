package util.tingfeng.java.files.officeToPdf;

public class OfficeToPdfFactory {
	private static OfficeToPdfI officeToPDFI=null;
	public enum OfficeType{
		OpenOffice,MsOffice,Wps
	}
	public synchronized static OfficeToPdfI getOfficeToPdfTools(OfficeType officeType){
		if(officeToPDFI!=null&&!officeType.equals(officeToPDFI.getOfficeType())){
			officeToPDFI=null;
		}
	switch (officeType) {
		case OpenOffice:{
			 if(officeToPDFI==null)
			 {
				 officeToPDFI=new OfficeToPdfByOpenOffice();
			 }
		break;
		}
		case MsOffice:
		{
			 if(officeToPDFI==null)
			 {
				 officeToPDFI=new OfficeToPdfbyMsOffice();
			 }
			break;
		}
		case Wps:
		{
			 if(officeToPDFI==null)
			 {
				 officeToPDFI=new OfficeToPdfByWps();
			 }
			 break;
		}
		default:
			{
				if(officeToPDFI==null)
				 {
					 officeToPDFI=new OfficeToPdfbyMsOffice();
				 }
				break;	
			}
		}
		return officeToPDFI;
	}
	
}
