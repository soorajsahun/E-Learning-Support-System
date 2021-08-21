package hit.PdfECertificateProject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfAction;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfDestination;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.TextField;

public class PdfECertificate {
	public static void main(String[] args)throws Exception {
		
		// Enter a valid Path with FileName
		String pdfName="E:\\PdfECertificateCreation\\pdfECertificate.pdf";
		
		// Create a empty PDF
		createEmptyCertificatePdf();
		
		// Add Image & Content to Certificate 
		addContentToCertificate(pdfName);
	}
	public static void createEmptyCertificatePdf() throws Exception{
		String temp="E:\\PdfECertificateCreation\\temp.pdf";
		try {
		//steps to create pdf
		// 1. Create document
		Document document=new Document(PageSize.A4);
		
		// 2. Create PdfWriter
		PdfWriter.getInstance(document, new FileOutputStream(temp));
		
		// 3. Open document
		document.open();
		
		// 4. Add content
		Paragraph para=new Paragraph(" ");
		document.add(para);
		
		// 5. Close document
		document.close();
		System.out.println("Empty Pdf Generated");
		}
		catch (Exception e) {
			System.out.println(e);
		}
	}
	public static void addContentToCertificate(String pdfName)throws Exception {
		Scanner input=new Scanner(System.in);
		try {
		//to add content in existing empty pdf-temp.pdf
		// Create Reader Instance
		PdfReader pdfReader=new PdfReader("temp.pdf");//Reads a PDF document.
		
		// Create Stamper Instance
		PdfStamper pdfStamper=new PdfStamper(pdfReader, new FileOutputStream(pdfName));
		//pdfStamper=Applies extra content to the pages of a PDF document
		//FileOutputStream=A file output stream is an output stream for writing data to a File
		
		// Set the Default Zoom-75%
		PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, pdfReader.getPageSize(1).getHeight(),
				1f);//A PdfDestination is a reference to a location in a PDF file and set the default size
		
		PdfAction action = PdfAction.gotoLocalPage(1, pdfDest, pdfStamper.getWriter());
//		PdfAction=Launches an application or a document->the application to be launched or the document to be opened or printed.
		
		pdfStamper.getWriter().setOpenAction(action);//Inserting attachment in PDF file using Java and iText
		
		// Create the Image Instance
		Image image = Image.getInstance("E:\\PdfECertificateCreation\\1.jpg");
		
		for (int i = 1; i <= pdfReader.getNumberOfPages(); i++) {
			if(i==1) {
				// -------------------- Background Image ----------------
				// Add Background Image (put content under)
				PdfContentByte content = pdfStamper.getUnderContent(i);
				image.setAbsolutePosition(0, 0);
				image.scaleToFit(PageSize.A4.getWidth(), PageSize.A4.getHeight());
				content.addImage(image);

				// --------------- Header Data ---------------------
				String headerData = "Course Completion Certificate\n\nThis is to certify that";
				// Enable During Final Test (To get input from User)
//					System.out.print("Enter Header Data : ");
//					headerData = input.nextLine();
				setData(pdfStamper, headerData, BaseFont.TIMES_BOLDITALIC, 30, BaseColor.BLACK, 530, 600, 70,
						450);

				// ---------------Main Content ----------------------
				// Add Student Name
				String studentName = "Mr. Name Surname ";
//					System.out.print("Enter Student Name on Certificate (FirstName SurnNme): ");
//					studentName = input.nextLine();
				
//				fetching email address
				File file=new File("E:\\PdfECertificateCreation\\ProgressReport.xls");
				FileInputStream fis=new FileInputStream(file);
				HSSFWorkbook workbook=new HSSFWorkbook(fis);
				HSSFSheet sheet=workbook.getSheetAt(0);
				String cellValue11=sheet.getRow(1).getCell(1).getStringCellValue();
				studentName = "Mr. "+cellValue11;
				
				setData(pdfStamper, studentName, BaseFont.TIMES_BOLDITALIC, 40, BaseColor.ORANGE, 530, 476, 80, 430);

				// ---------Middle Data (Reason of Certificate) ------------------
				// Prize Details & Standard Format of Data
				String midData = "Course on: JAVA FULL STACK DEVELOPMENT\n\n"
						+ "Successfylly Completed the COURSE on-time"
						+ " and his participation was very good. "
						+ "\nWe wish him good luck for his future.....";
				// Enable During Final Test (To get input from User)
//					System.out.print("Enter the Main Content: : ");
//					midData = input.nextLine();

				setData(pdfStamper, midData, BaseFont.TIMES_BOLDITALIC, 14, BaseColor.BLACK, 530, 290, 70, 410);

				// ----------Footer Data -------------------

				// Add Current Date
				Date date = new Date();
				setData(pdfStamper, date.toString(), BaseFont.TIMES_BOLDITALIC, 16, BaseColor.RED, 230, 220, 80,
						150);
				setData(pdfStamper, "Acquired on", BaseFont.TIMES_BOLDITALIC, 14, BaseColor.DARK_GRAY, 230, 170, 80,
						145);

				// Add Name of the Head Person
				String nameOfHeadPerson = "Mr. Name Surname";
				String direc = "Director of Haaris Infotech Pvt. Ltd.";
				// Enable During Final Test (To get input from User)
//					System.out.print("Enter Head/Director Name : ");
//					nameOfHeadPerson = input.nextLine();
//					nameOfHeadPerson = "Mr. "+nameOfHeadPerson;
				setData(pdfStamper, nameOfHeadPerson, BaseFont.TIMES_BOLDITALIC, 18, BaseColor.BLUE, 345, 220, 550,
						130);
				setData(pdfStamper, direc, BaseFont.TIMES_BOLDITALIC, 12, BaseColor.DARK_GRAY, 345, 170, 550, 125);
				
				fis.close();
				workbook.close();
			}
			input.close();
		}

		pdfStamper.close();
		pdfReader.close();

		System.out.println("Successfully added content & created the PDF E-Certificate : " + pdfName);

	} catch (IOException e) {
		System.out.println("Exception while adding Content To Certificate");
		e.printStackTrace();
	}
		
	}
	private static void setData(PdfStamper pdfStamper, String data,String fontType, float fontSize,
			BaseColor fontColor, float lowerX, float lowerY, float upperX, float upperY) throws Exception {
		try {
//
			TextField textField = new TextField(pdfStamper.getWriter(), new Rectangle(lowerX, lowerY, upperX, upperY),
					"newTextField");
			textField.setOptions(TextField.MULTILINE | TextField.READ_ONLY);
			textField.setAlignment(Element.ALIGN_CENTER);
			textField.setTextColor(fontColor);
			BaseFont baseFont = BaseFont.createFont(fontType, BaseFont.WINANSI, BaseFont.EMBEDDED);
			textField.setFont(baseFont);
			textField.setFontSize(fontSize);
			textField.setText(data);
//			// is no longer multiple-line
			pdfStamper.addAnnotation(textField.getTextField(), 1);

		} catch (Exception ex) {
			System.out.println("Exception while Setting Data: " + data);
			ex.printStackTrace();
		}
		
	}
	private static String getDataFromUser(String dataType) {
		String data = "";
		Scanner input = new Scanner(System.in);
		System.out.println("Enter " + dataType);
		data = input.nextLine();
		input.close();
		return data;
	}

}

