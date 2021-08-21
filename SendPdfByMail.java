package hit.Project;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Properties;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeSet;

import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class SendPdfByMail {
	public static void main(String[] args) throws Exception {
		String subject ="Course completion certificate!!!";
		String textBody = "<h1>CONGRATULATIONS!!!.<h1> <br/> <h2>You have successfully completed the \"JAVA FULLSTACK DEVELOPMENT\" course Sincerely.<h2> <br/>"
				+ "Your certificate is attached with this mail.<br/>"
				+ "\nThanks & Regards,<br/>Suraj Sahu.";
	
		String myEmailAccount="suraj.sahu.9484@gmail.com";
		String password="9819675309";
		sendMailFromGmail(subject, textBody, myEmailAccount, password);
	}
	public static void sendMailFromGmail(String subject,String textBody,final String myEmailAccount,final String password) throws Exception {
//		fetching email address
		File file=new File("E:\\PdfECertificateCreation\\ProgressReport.xls");
		FileInputStream fis=new FileInputStream(file);
		HSSFWorkbook workbook=new HSSFWorkbook(fis);
		HSSFSheet sheet=workbook.getSheetAt(0);
		String cellValue12=sheet.getRow(1).getCell(2).getStringCellValue();
//		String cellValue22=sheet.getRow(2).getCell(2).getStringCellValue();
		
		//recipients email addresses
		Set<String> recipients=new TreeSet<String>();
			recipients.add(cellValue12);
//			recipients.add(cellValue22);
			
		//creating session
		Properties prop = new Properties();
		prop.put("mail.smtp.host", "smtp.gmail.com");
		prop.put("mail.smtp.port","465");
		prop.put("mail.smtp.ssl.enable","true");
		prop.put("mail.smtp.auth","true");
		
		Session session=Session.getInstance(prop, new Authenticator() {//session object stores all the information of host like host name, username, password etc.
			@Override
			protected PasswordAuthentication getPasswordAuthentication() {
				// TODO Auto-generated method stub
				return new PasswordAuthentication(myEmailAccount, password);
			}
		});
		
		//Compose message
		try {
		MimeMessage message=new MimeMessage(session);
		
		message.setFrom(new InternetAddress(myEmailAccount));
		for(String recipient:recipients) {
			if(isValidEmailAddress(recipient)==true) {
				String userName=getUserName(recipient);
				message.addRecipient(Message.RecipientType.TO, new InternetAddress(recipient));
				
				MimeMultipart mimeMultiPart=new MimeMultipart();
				
				String path1="E:\\PdfECertificateCreation\\pdfECertificate.pdf";
				
				MimeBodyPart t1=new MimeBodyPart();
				MimeBodyPart a1=new MimeBodyPart();
				
				message.setSubject(subject);
				t1.setContent(textBody, "text/html");
				
				File file1=new File(path1);

				try {
				a1.attachFile(file1);
				
				} catch (Exception e) {
					e.printStackTrace();
				}
				
				mimeMultiPart.addBodyPart(t1);
				mimeMultiPart.addBodyPart(a1);
				message.setContent(mimeMultiPart);
				
				System.out.println("Sending the mail...");
				Transport.send(message);
				System.out.println("Mail sent succesfully...");
			}
		}
		
		}catch(MessagingException mex) {
			mex.printStackTrace();
		}
		fis.close();
		workbook.close();
	}
	//checking valid email
	public static boolean isValidEmailAddress(String email) {
		   boolean result = true;
		   try {
		      InternetAddress emailAddr = new InternetAddress(email);
		      emailAddr.validate();
		   } catch (AddressException ex) {
		      result = false;
		   }
		   return result;
		}
	
	//extracting user name from email
	public static String getUserName(String email) {
		int index=email.indexOf("@");
		email=email.substring(0, index);
		return email;
	}
}



