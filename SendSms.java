package hit.Project;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.twilio.Twilio;
import com.twilio.rest.api.v2010.account.Message;

public class SendSms {
	static int otp;
	//Enter your Twilio ACCOUNT_SID & AUTH_TOKEN
	public static final String ACCOUNT_SID ="ACebb31a3f3f3106b84bf0f5380edbae08";
	public static final String AUTH_TOKEN ="046b518727f180c2da0532099bb30a58";
	
	public static void sendSms() throws Exception {
		//fetching contact no
		File file=new File("E:\\PdfECertificateCreation\\ProgressReport.xls");
		FileInputStream fis=new FileInputStream(file);
		HSSFWorkbook workbook=new HSSFWorkbook(fis);
		HSSFSheet sheet=workbook.getSheetAt(0);
		String cellValue13=sheet.getRow(1).getCell(3).getStringCellValue();
//		String cellValue23=sheet.getRow(2).getCell(3).getStringCellValue();
		
		ArrayList<String> phoneNumbersList=new ArrayList<String>();
			phoneNumbersList.add(cellValue13);
//			phoneNumbersList.add(cellValue23);
			
//		Message
		String cellValue11=sheet.getRow(1).getCell(1).getStringCellValue();
		String sb="Hi "+cellValue11+"!\n"+"Congratulations!\n"+"You'hv successfully completed registration for \"JAVA FULLSTACK DEVELOPMENT\"."
				+ "Your registration is confirmed."
				+ "Kindly Check your mail for further details. ";
		
		try {
			Twilio.init(ACCOUNT_SID, AUTH_TOKEN);
			
			for (String phoneNumber : phoneNumbersList) {
				//In Message include the to Phone Number, Message & From Phone Number(Twilio account)
				Message message = Message.creator(new com.twilio.type.PhoneNumber("+91" + phoneNumber),
						new com.twilio.type.PhoneNumber("+1 772 320 5742"), "\n"+sb).create();
				//Get the Transaction Id
				System.out.println("Session ID/String Identifier (SID): "+message.getSid());
				System.out.println("SMS Sent Successfully to :"+phoneNumber);
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		fis.close();
		workbook.close();
	}
}


