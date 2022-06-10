package XLXM.xlxmcopy;

import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

public class JdbcMsSql {
	public static String username;
	public static String pass;
	public static Connection connObj;
	public static String db;
	public static String summary = "Z_Sales_Written_Manager_SummaryTable"; //for summary table name
	public static String analysis = "Z_Sales_Written_Manager_Principal_Product_Qty"; // for analysis table name
	public static String summary2 = "Z_Sales_Written_Manager_Principal_Particular_Product_Qty"; //for summary table name
	public static String analysis2 = "Z_Sales_Written_Manager_Principal_Particular_Product_Val"; // for analysis table name
	

	public static List<ProductAnalysis> getDbConnection() {
		String JDBC_URL = "jdbc:sqlserver://localhost;user="+username+";password="+pass+";";
		List<ProductAnalysis> productList = new ArrayList<ProductAnalysis>();
		try {
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			connObj = DriverManager.getConnection(JDBC_URL);
			if (connObj != null) {
				String SQL = "SELECT * FROM "+db+".dbo."+analysis+";";
				Statement stmt = connObj.createStatement();
				ResultSet rs = stmt.executeQuery(SQL);
				while (rs.next()) {
					ProductAnalysis product = new ProductAnalysis();
			
					product.Manager=rs.getString("Manager");
					product.Principal=rs.getString("Principal");
					product.Product=rs.getString("Product");
					product.MontylyTargetQty=rs.getLong("MontylyTargetQty");
					product.PricePerUnit=rs.getLong("PricePerUnit");
					product.MonthlyTargetVal=rs.getLong("MonthlyTargetVal");
					product.Sales2019=rs.getLong("2019_Sales");
					product.Sales2020=rs.getLong("2020_Sales");
					product.Jan=rs.getLong("Jan");
					product.Feb=rs.getLong("Feb");
					product.Mar=rs.getLong("Mar");
					product.Apr=rs.getLong("Apr");
					product.May=rs.getLong("May");
					product.Jun=rs.getLong("Jun");
					product.Jul=rs.getLong("Jul");
					product.Aug=rs.getLong("Aug");
					product.Sep=rs.getLong("Sep");
					product.Oct=rs.getLong("Oct");
					product.Nov=rs.getLong("Nov");
					product.Dec=rs.getLong("Dec");
					product.GrandTotal=rs.getLong("GrandTotal");
					product.TargetTillLastMonth=rs.getLong("TargetTillLastMonth");
					product.SalesTillLastMonth=rs.getLong("SalesTillLastMonth");
					product.Revised_Target=rs.getLong("Revised_Target");
					product.Jan2020Qty=rs.getLong("Jan2020Qty");
					product.Feb2020Qty=rs.getLong("Feb2020Qty");
					product.Mar2020Qty=rs.getLong("Mar2020Qty");
					product.Apr2020Qty=rs.getLong("Apr2020Qty");
					product.May2020Qty=rs.getLong("May2020Qty");
					product.Jun2020Qty=rs.getLong("Jun2020Qty");
					product.Jul2020Qty=rs.getLong("Jul2020Qty");
					product.Aug2020Qty=rs.getLong("Aug2020Qty");
					product.Sep2020Qty=rs.getLong("Sep2020Qty");
					product.Oct2020Qty=rs.getLong("Oct2020Qty");
					product.Nov2020Qty=rs.getLong("Nov2020Qty");
					product.Dec2020Qty=rs.getLong("Dec2020Qty");
					product.GrandTotalQty=rs.getLong("GrandTotalQty");

		
					
					// you can add columns here like line above just check datatype
					productList.add(product);
					//System.out.println(rs.getString("Principal") + " : " + rs.getString("Location"));
				}
			}
		} catch (Exception sqlException) {
			sqlException.printStackTrace();
		}
		return productList;
	}
	
	
	
	//---------------------------------------//
	
	public static List<ProductAnalysis2> getDbConnection1() {
		String JDBC_URL = "jdbc:sqlserver://localhost;user="+username+";password="+pass+";";
		List<ProductAnalysis2> productList1 = new ArrayList<ProductAnalysis2>();
		try {
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			connObj = DriverManager.getConnection(JDBC_URL);
			if (connObj != null) {
				String SQL = "SELECT * FROM "+db+".dbo."+analysis2+";";
				Statement stmt = connObj.createStatement();
				ResultSet rs = stmt.executeQuery(SQL);
				while (rs.next()) {
					ProductAnalysis2 product2 = new ProductAnalysis2();
			
						product2.Manager=rs.getString("Manager");
						product2.Particulars=rs.getString("Particulars");
						product2.Principal=rs.getString("Principal");
						product2.Product=rs.getString("Product");
						product2.Location=rs.getString("Location");
						product2.GrandTotalVal=rs.getLong("GrandTotalVal");
						product2.JanVal=rs.getLong("JanVal");
						product2.FebVal=rs.getLong("FebVal");
						product2.MarVal=rs.getLong("MarVal");
						product2.AprVal=rs.getLong("AprVal");
						product2.MayVal=rs.getLong("MayVal");
						product2.JunVal=rs.getLong("JunVal");
						product2.JulVal=rs.getLong("JulVal");
						product2.AugVal=rs.getLong("AugVal");
						product2.SepVal=rs.getLong("SepVal");
						product2.OctVal=rs.getLong("OctVal");
						product2.NovVal=rs.getLong("NovVal");
						product2.DecVal=rs.getLong("DecVal");
						product2.GrandTotalVal020=rs.getLong("GrandTotalVal020");
						product2.Jan=rs.getLong("Jan");
						product2.Feb=rs.getLong("Feb");
						product2.Mar=rs.getLong("Mar");
						product2.Apr=rs.getLong("Apr");
						product2.May=rs.getLong("May");
						product2.Jun=rs.getLong("Jun");
						product2.Jul=rs.getLong("Jul");
						product2.Aug=rs.getLong("Aug");
						product2.Sep=rs.getLong("Sep");
						product2.Oct=rs.getLong("Oct");
						product2.Nov=rs.getLong("Nov");
						product2.Dec=rs.getLong("Dec");
						product2.GrandTotalVal2021=rs.getLong("GrandTotalVal2021");


		
					
					// you can add columns here like line above just check datatype
					productList1.add(product2);
					//System.out.println(rs.getString("Principal") + " : " + rs.getString("Location"));
				}
			}
		} catch (Exception sqlException) {
			sqlException.printStackTrace();
		}
		return productList1;
	}
	
	
	//------------------------------------//
	
	
	public static List<Summary> getSummary() {
		String JDBC_URL = "jdbc:sqlserver://localhost;user="+username+";password="+pass+";";
		List<Summary> summaryList = new ArrayList<Summary>();
		try {
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			connObj = DriverManager.getConnection(JDBC_URL);
			if (connObj != null) {
				String SQL = "SELECT * FROM "+db+".dbo."+summary+";";
				Statement stmt = connObj.createStatement();
				ResultSet rs = stmt.executeQuery(SQL);
				while (rs.next()) {
					Summary summary = new Summary();
		
					summary.Manager=rs.getString("Manager");
					summary.Principal=rs.getString("Principal");
					summary.Target=rs.getLong("Target");
					summary.Sales2019=rs.getLong("2019Sales");
					summary.Sales2020=rs.getLong("2020Sales");
					summary.Jan_Val=rs.getLong("Jan_Val");
					summary.Feb_Val=rs.getLong("Feb_Val");
					summary.Mar_Val=rs.getLong("Mar_Val");
					summary.Apr_Val=rs.getLong("Apr_Val");
					summary.May_Val=rs.getLong("May_Val");
					summary.Jun_Val=rs.getLong("Jun_Val");
					summary.Jul_Val=rs.getLong("Jul_Val");
					summary.Aug_Val=rs.getLong("Aug_Val");
					summary.Sep_Val=rs.getLong("Sep_Val");
					summary.Oct_Val=rs.getLong("Oct_Val");
					summary.Nov_Val=rs.getLong("Nov_Val");
					summary.Dec_Val=rs.getLong("Dec_Val");
					summary.TargettilllastMonth=rs.getLong("TargettilllastMonth");
					summary.SalestilllastMonth=rs.getLong("SalestilllastMonth");
					summary.Achievement=rs.getLong("Achievement");
					summary.GrandTotalVal=rs.getLong("GrandTotalVal");
					summary.Revised_Target=rs.getLong("Revised_Target");


					//for summary you can add here
					summaryList.add(summary);				}
			}
		} catch (Exception sqlException) {
			sqlException.printStackTrace();
		}
		return summaryList;
	}
	
	
	
	//---------------------- Value Wise------------------------//
	
	public static List<Summary2> getSummary2() {
		String JDBC_URL1 = "jdbc:sqlserver://localhost;user="+username+";password="+pass+";";
		List<Summary2> summaryList2 = new ArrayList<Summary2>();
		try {
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			connObj = DriverManager.getConnection(JDBC_URL1);
			if (connObj != null) {
				String SQL = "SELECT * FROM "+db+".dbo."+summary2+";";
				Statement stmt = connObj.createStatement();
				ResultSet rs = stmt.executeQuery(SQL);
				while (rs.next()) {
					Summary2 summary2 = new Summary2();
		
					
					summary2.Manager=rs.getString("Manager");
					summary2.Particulars=rs.getString("Particulars");
					summary2.Principal=rs.getString("Principal");
					summary2.Product=rs.getString("Product");
					summary2.GrandTotalQty=rs.getString("GrandTotalQty");
					summary2.JanQty=rs.getLong("JanQty");
					summary2.FebQty=rs.getLong("FebQty");
					summary2.MarQty=rs.getLong("MarQty");
					summary2.AprQty=rs.getLong("AprQty");
					summary2.MayQty=rs.getLong("MayQty");
					summary2.JunQty=rs.getLong("JunQty");
					summary2.JulQty=rs.getLong("JulQty");
					summary2.AugQty=rs.getLong("AugQty");
					summary2.SepQty=rs.getLong("SepQty");
					summary2.OctQty=rs.getLong("OctQty");
					summary2.NovQty=rs.getLong("NovQty");
					summary2.DecQty=rs.getLong("DecQty");
					summary2.GrandTotalQty2020=rs.getLong("GrandTotalQty2020");
					summary2.Jan=rs.getLong("Jan");
					summary2.Feb=rs.getLong("Feb");
					summary2.Mar=rs.getLong("Mar");
					summary2.Apr=rs.getLong("Apr");
					summary2.May=rs.getLong("May");
					summary2.Jun=rs.getLong("Jun");
					summary2.Jul=rs.getLong("Jul");
					summary2.Aug=rs.getLong("Aug");
					summary2.Sep=rs.getLong("Sep");
					summary2.Oct=rs.getLong("Oct");
					summary2.Nov=rs.getLong("Nov");
					summary2.Dec=rs.getLong("Dec");
					summary2.GrandTotalQty2021=rs.getLong("GrandTotalQty2021");
					summary2.Jan21Target=rs.getLong("Jan21Target");
					summary2.Feb21Target=rs.getLong("Feb21Target");
					summary2.Mar21Target=rs.getLong("Mar21Target");
					summary2.Apr21Target=rs.getLong("Apr21Target");
					summary2.May21Target=rs.getLong("May21Target");
					summary2.Jun21Target=rs.getLong("Jun21Target");
					summary2.Jul21Target=rs.getLong("Jul21Target");
					summary2.Aug21Target=rs.getLong("Aug21Target");
					summary2.Sep21Target=rs.getLong("Sep21Target");
					summary2.Oct21Target=rs.getLong("Oct21Target");
					summary2.Nov21Target=rs.getLong("Nov21Target");
					summary2.Dec21Target=rs.getLong("Dec21Target");


					
					//for summary you can add here
					summaryList2.add(summary2);				}
			}
		} catch (Exception sqlException) {
			sqlException.printStackTrace();
		}
		return summaryList2;
	}
	
}