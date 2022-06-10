package XLXM.xlxmcopy;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteToFile {
	public static void main(String[] args) throws IOException {
		String temp = "D:\\PPZL_SalesReport_Manager_Summy.xlsm";
		String myFile = args.length>0?args[0]:temp;
		JdbcMsSql.username = args.length>1?args[1]:"SERVERNNAME";
		JdbcMsSql.pass = args.length>2?args[2]:"PASSWORDOFTHEUSER";
		JdbcMsSql.db = args.length>3?args[3]:"DATABASENAME";
		if(args.length > 4)
			JdbcMsSql.analysis = args[4];

		if(args.length > 5)
			JdbcMsSql.summary = args[5];
		if(args.length > 6)
			JdbcMsSql.summary2 = args[6];
		
		FileInputStream file = new FileInputStream(new File(myFile));
		System.out.println("found file");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		System.out.println("in workbook");
		
		
		XSSFSheet sheet = workbook.getSheet("PrincipalProductSummaryTable");
		List<ProductAnalysis> productList = JdbcMsSql.getDbConnection();
		System.out.println(sheet.getLastRowNum());
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			XSSFRow row = sheet.getRow(i);
			if (row != null)
				deleteRow(sheet, row);
		}
		int last = 1;
		for (ProductAnalysis productAnalysis : productList) {
			XSSFRow row1 = sheet.createRow(last++);
	
			row1.createCell(0).setCellValue(productAnalysis.Manager);
			row1.createCell(1).setCellValue(productAnalysis.Principal);
			row1.createCell(2).setCellValue(productAnalysis.Product);
			row1.createCell(3).setCellValue(productAnalysis.MontylyTargetQty);
			row1.createCell(4).setCellValue(productAnalysis.PricePerUnit);
			row1.createCell(5).setCellValue(productAnalysis.MonthlyTargetVal);
			row1.createCell(6).setCellValue(productAnalysis.Sales2019);
			row1.createCell(7).setCellValue(productAnalysis.Sales2020);
			row1.createCell(8).setCellValue(productAnalysis.Jan);
			row1.createCell(9).setCellValue(productAnalysis.Feb);
			row1.createCell(10).setCellValue(productAnalysis.Mar);
			row1.createCell(11).setCellValue(productAnalysis.Apr);
			row1.createCell(12).setCellValue(productAnalysis.May);
			row1.createCell(13).setCellValue(productAnalysis.Jun);
			row1.createCell(14).setCellValue(productAnalysis.Jul);
			row1.createCell(15).setCellValue(productAnalysis.Aug);
			row1.createCell(16).setCellValue(productAnalysis.Sep);
			row1.createCell(17).setCellValue(productAnalysis.Oct);
			row1.createCell(18).setCellValue(productAnalysis.Nov);
			row1.createCell(19).setCellValue(productAnalysis.Dec);
			row1.createCell(20).setCellValue(productAnalysis.GrandTotal);
			row1.createCell(21).setCellValue(productAnalysis.TargetTillLastMonth);
			row1.createCell(22).setCellValue(productAnalysis.SalesTillLastMonth);
			row1.createCell(23).setCellValue(productAnalysis.Revised_Target);
			row1.createCell(24).setCellValue(productAnalysis.Jan2020Qty);
			row1.createCell(25).setCellValue(productAnalysis.Feb2020Qty);
			row1.createCell(26).setCellValue(productAnalysis.Mar2020Qty);
			row1.createCell(27).setCellValue(productAnalysis.Apr2020Qty);
			row1.createCell(28).setCellValue(productAnalysis.May2020Qty);
			row1.createCell(29).setCellValue(productAnalysis.Jun2020Qty);
			row1.createCell(30).setCellValue(productAnalysis.Jul2020Qty);
			row1.createCell(31).setCellValue(productAnalysis.Aug2020Qty);
			row1.createCell(32).setCellValue(productAnalysis.Sep2020Qty);
			row1.createCell(33).setCellValue(productAnalysis.Oct2020Qty);
			row1.createCell(34).setCellValue(productAnalysis.Nov2020Qty);
			row1.createCell(35).setCellValue(productAnalysis.Dec2020Qty);
			row1.createCell(36).setCellValue(productAnalysis.GrandTotalQty);

			
			
			//you can add columns in xlx from here just keep the order
		}

		
		////   --------------------------//
		
		XSSFSheet sheetp = workbook.getSheet("PrincipalParticularVal");
		List<ProductAnalysis2> productList1 = JdbcMsSql.getDbConnection1();
		System.out.println(sheetp.getLastRowNum());
		for (int i = 1; i <= sheetp.getLastRowNum(); i++) {
			XSSFRow row = sheetp.getRow(i);
			if (row != null)
				deleteRow(sheetp, row);
		}
		int last11 = 1;
		for (ProductAnalysis2 productAnalysis2 : productList1) {
			XSSFRow row1 = sheetp.createRow(last11++);
	
				row1.createCell(0).setCellValue(productAnalysis2.Manager);
				row1.createCell(1).setCellValue(productAnalysis2.Particulars);
				row1.createCell(2).setCellValue(productAnalysis2.Principal);
				row1.createCell(3).setCellValue(productAnalysis2.Product);
				row1.createCell(4).setCellValue(productAnalysis2.Location);
				row1.createCell(5).setCellValue(productAnalysis2.GrandTotalVal);
				row1.createCell(6).setCellValue(productAnalysis2.JanVal);
				row1.createCell(7).setCellValue(productAnalysis2.FebVal);
				row1.createCell(8).setCellValue(productAnalysis2.MarVal);
				row1.createCell(9).setCellValue(productAnalysis2.AprVal);
				row1.createCell(10).setCellValue(productAnalysis2.MayVal);
				row1.createCell(11).setCellValue(productAnalysis2.JunVal);
				row1.createCell(12).setCellValue(productAnalysis2.JulVal);
				row1.createCell(13).setCellValue(productAnalysis2.AugVal);
				row1.createCell(14).setCellValue(productAnalysis2.SepVal);
				row1.createCell(15).setCellValue(productAnalysis2.OctVal);
				row1.createCell(16).setCellValue(productAnalysis2.NovVal);
				row1.createCell(17).setCellValue(productAnalysis2.DecVal);
				row1.createCell(18).setCellValue(productAnalysis2.GrandTotalVal020);
				row1.createCell(19).setCellValue(productAnalysis2.Jan);
				row1.createCell(20).setCellValue(productAnalysis2.Feb);
				row1.createCell(21).setCellValue(productAnalysis2.Mar);
				row1.createCell(22).setCellValue(productAnalysis2.Apr);
				row1.createCell(23).setCellValue(productAnalysis2.May);
				row1.createCell(24).setCellValue(productAnalysis2.Jun);
				row1.createCell(25).setCellValue(productAnalysis2.Jul);
				row1.createCell(26).setCellValue(productAnalysis2.Aug);
				row1.createCell(27).setCellValue(productAnalysis2.Sep);
				row1.createCell(28).setCellValue(productAnalysis2.Oct);
				row1.createCell(29).setCellValue(productAnalysis2.Nov);
				row1.createCell(30).setCellValue(productAnalysis2.Dec);
				row1.createCell(31).setCellValue(productAnalysis2.GrandTotalVal2021);


			
			
			//you can add columns in xlx from here just keep the order
		}

		
		
		//-----------------------------//
		
		
		
		
		XSSFSheet analysisSheet = workbook.getSheet("PrincipalSummaryTable");
		List<Summary> summaryList = JdbcMsSql.getSummary();
		for (int i = 1; i <= analysisSheet.getLastRowNum(); i++) {
			XSSFRow row = analysisSheet.getRow(i);
			if (row != null)
				deleteRow(analysisSheet, row);
		}
		int lastAnalysis = 1;
		for (Summary summary : summaryList) {
			XSSFRow row1 = analysisSheet.createRow(lastAnalysis++);
			
			row1.createCell(0).setCellValue(summary.Manager);
			row1.createCell(1).setCellValue(summary.Principal);
			row1.createCell(2).setCellValue(summary.Target);
			row1.createCell(3).setCellValue(summary.Sales2019);
			row1.createCell(4).setCellValue(summary.Sales2020);
			row1.createCell(5).setCellValue(summary.Jan_Val);
			row1.createCell(6).setCellValue(summary.Feb_Val);
			row1.createCell(7).setCellValue(summary.Mar_Val);
			row1.createCell(8).setCellValue(summary.Apr_Val);
			row1.createCell(9).setCellValue(summary.May_Val);
			row1.createCell(10).setCellValue(summary.Jun_Val);
			row1.createCell(11).setCellValue(summary.Jul_Val);
			row1.createCell(12).setCellValue(summary.Aug_Val);
			row1.createCell(13).setCellValue(summary.Sep_Val);
			row1.createCell(14).setCellValue(summary.Oct_Val);
			row1.createCell(15).setCellValue(summary.Nov_Val);
			row1.createCell(16).setCellValue(summary.Dec_Val);
			row1.createCell(17).setCellValue(summary.TargettilllastMonth);
			row1.createCell(18).setCellValue(summary.SalestilllastMonth);
			row1.createCell(19).setCellValue(summary.Achievement);
			row1.createCell(20).setCellValue(summary.GrandTotalVal);
			row1.createCell(21).setCellValue(summary.Revised_Target);

			
			
			//you can add columns in xlx from here just keep the order
			//done
		}
	
		
		
		//------------------------New One Value wise---
		XSSFSheet sheet1 = workbook.getSheet("CustomerQtyTable");
		List<Summary2> summaryList2 = JdbcMsSql.getSummary2();
	//	System.out.println(sheet1.getLastRowNum());
		for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
			XSSFRow row = sheet1.getRow(i);
			if (row != null)
				deleteRow(sheet1, row);
		}
		int last1 = 1;
		for (Summary2 summary2: summaryList2) {
			XSSFRow row1 = sheet1.createRow(last1++);
	
			row1.createCell(0).setCellValue(summary2.Manager);
			row1.createCell(1).setCellValue(summary2.Particulars);
			row1.createCell(2).setCellValue(summary2.Principal);
			row1.createCell(3).setCellValue(summary2.Product);
			row1.createCell(4).setCellValue(summary2.GrandTotalQty);
			row1.createCell(5).setCellValue(summary2.JanQty);
			row1.createCell(6).setCellValue(summary2.FebQty);
			row1.createCell(7).setCellValue(summary2.MarQty);
			row1.createCell(8).setCellValue(summary2.AprQty);
			row1.createCell(9).setCellValue(summary2.MayQty);
			row1.createCell(10).setCellValue(summary2.JunQty);
			row1.createCell(11).setCellValue(summary2.JulQty);
			row1.createCell(12).setCellValue(summary2.AugQty);
			row1.createCell(13).setCellValue(summary2.SepQty);
			row1.createCell(14).setCellValue(summary2.OctQty);
			row1.createCell(15).setCellValue(summary2.NovQty);
			row1.createCell(16).setCellValue(summary2.DecQty);
			row1.createCell(17).setCellValue(summary2.GrandTotalQty2020);
			row1.createCell(18).setCellValue(summary2.Jan);
			row1.createCell(19).setCellValue(summary2.Feb);
			row1.createCell(20).setCellValue(summary2.Mar);
			row1.createCell(21).setCellValue(summary2.Apr);
			row1.createCell(22).setCellValue(summary2.May);
			row1.createCell(23).setCellValue(summary2.Jun);
			row1.createCell(24).setCellValue(summary2.Jul);
			row1.createCell(25).setCellValue(summary2.Aug);
			row1.createCell(26).setCellValue(summary2.Sep);
			row1.createCell(27).setCellValue(summary2.Oct);
			row1.createCell(28).setCellValue(summary2.Nov);
			row1.createCell(29).setCellValue(summary2.Dec);
			row1.createCell(30).setCellValue(summary2.GrandTotalQty2021);
			row1.createCell(31).setCellValue(summary2.Jan21Target);
			row1.createCell(32).setCellValue(summary2.Feb21Target);
			row1.createCell(33).setCellValue(summary2.Mar21Target);
			row1.createCell(34).setCellValue(summary2.Apr21Target);
			row1.createCell(35).setCellValue(summary2.May21Target);
			row1.createCell(36).setCellValue(summary2.Jun21Target);
			row1.createCell(37).setCellValue(summary2.Jul21Target);
			row1.createCell(38).setCellValue(summary2.Aug21Target);
			row1.createCell(39).setCellValue(summary2.Sep21Target);
			row1.createCell(40).setCellValue(summary2.Oct21Target);
			row1.createCell(41).setCellValue(summary2.Nov21Target);
			row1.createCell(42).setCellValue(summary2.Dec21Target);


			//row1.createCell(11).setCellValue(productAnalysis.SalesAccount);
			
			
			//you can add columns in xlx from here just keep the order
		}

		
		
		
		
		
		//---------------Write Value wise------------------//
		
	
	
		
		FileOutputStream os = new FileOutputStream(myFile);
		workbook.write(os);
		System.out.println("Writing on XLSX file Finished ...");
		file.close();
	}

	public static void deleteRow(XSSFSheet sheet, XSSFRow row) {
		int lastRowNum = sheet.getLastRowNum();
		int rowIndex = row.getRowNum();
		if (rowIndex >= 0 && rowIndex < lastRowNum) {
			sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
		}
		if (rowIndex == lastRowNum) {
			XSSFRow removingRow = sheet.getRow(rowIndex);
			if (removingRow != null) {
				sheet.removeRow(removingRow);
				System.out.println("Deleting.... ");
			}
		}
	}
}
