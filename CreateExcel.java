package hit.Project;

import java.io.FileOutputStream;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.itextpdf.text.log.SysoLogger;

public class CreateExcel {
	public static void callCreateExcel() throws Exception{
		String fileName="E:\\PdfECertificateCreation\\ProgressReport.xls";
		createExcel(fileName);
	}
	public static String[] rowHeads;
	public static void createExcel(String fileName)throws Exception {
			System.out.println("-------- Creating an Excel Sheet ----------");
			System.out.println("Enter columns as username,email and contact");
//			Workbook
//			Sheet
//			Row
//			Cell

			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("StudentMarksReport");

			Scanner input = new Scanner(System.in);
			
			System.out.print("Enter Number of Coloumns: ");
			int numberOdColoumns = input.nextInt();
			
			System.out.print("Enter Number of Rows: ");
			int numberOfRows = input.nextInt();
			
			rowHeads = new String[numberOdColoumns];

			createRowHead(sheet, numberOdColoumns);

			for (int cnt = 0; cnt < numberOfRows; cnt++) {
				createRows(sheet, cnt);
			}
			
			FileOutputStream fileOut = new FileOutputStream(fileName);
			workbook.write(fileOut);
			
			fileOut.close();
			workbook.close();
			
			System.out.println("Successfully Created the Excel: "+fileName);
		
	}

	public static void createRows(HSSFSheet sheet, int rowNumber) throws Exception {
		String data = "";
		Scanner input = new Scanner(System.in);
			rowNumber = rowNumber+1;	
			HSSFRow row = sheet.createRow((short) rowNumber);	
			row.createCell(0).setCellValue(rowNumber);	
			for (int i = 0; i < rowHeads.length; i++) {
				System.out.print("Enter Data for " + rowHeads[i]+": ");
				data = input.nextLine();
				row.createCell(i+1).setCellValue(data);
			}
			System.out.println("Successfully Created row no: "+rowNumber);
		
	}

	public static void createRowHead(HSSFSheet sheet, int numberOdColoumns) throws Exception {
		Scanner input = new Scanner(System.in);
		
			HSSFRow row = sheet.createRow((short) 0);
			System.out.println("------- Note: S.No Coloumn was creating as By default --------");
			row.createCell(0).setCellValue("S.No");
			for (int i = 0; i < numberOdColoumns; i++) {
				System.out.print("Enter Data into Coloumn-" + (i+1)+" : ");
				rowHeads[i] = input.nextLine();
				row.createCell(i+1).setCellValue(rowHeads[i]);
			}
			System.out.println("Successfully Created Head row");
		
	}
}

