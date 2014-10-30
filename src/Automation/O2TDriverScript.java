/*################################################################# #####################################
'Author		       	: Open2Test
'Version	    	: V 1.2
'Date of Creation	: 5-JUL-2013
'#######################################################################################################
 */

package Automation;

import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

import org.apache.commons.io.FileUtils;
//import org.junit.After;
//import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverBackedSelenium;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.android.AndroidDriver;
import org.openqa.selenium.chrome.ChromeDriver;
//import org.openqa.selenium.android.library.AndroidWebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.firefox.internal.ProfilesIni;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.internal.seleniumemulation.IsElementPresent;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.sikuli.api.DesktopScreenRegion;
import org.sikuli.api.ImageTarget;
import org.sikuli.api.ScreenRegion;
import org.sikuli.api.Target;
import org.sikuli.api.robot.desktop.DesktopMouse;
import org.sikuli.api.visual.Canvas;
import org.sikuli.api.visual.DesktopCanvas;
import org.sikuli.script.*;
import org.testng.annotations.AfterTest;
import org.testng.annotations.Test;

import autoitx4java.AutoItX;

import com.jacob.com.LibraryLoader;

public class O2TDriverScript

{
	boolean boolval;
	public static String DR_URL;
	public static int TC_PASS=0;
	public static int TC_FAIL=0;
	public static String filenamer;
	static BufferedWriter bw = null;
	static BufferedWriter bw1 = null;
	static   WebDriver D8;
	static  WebDriverBackedSelenium  D9;
	Date cur_dt = null;
	String TestSuite;
	String TestScript;
	String ObjectRepository;
	String ReusableComponents;
	static String TestData;
	int startrow = 0;
	static String ReportsPath;
	static String TestSummaryReport;
	String strResultPath;
	String[] TCNm = null;
	static String exeStatus = "True";
	static String TestReport;
	static int rowcnt;
	static int dtrownum = 1;
	static int ORrowcount = 0;
	// int loopcount = 0;
	static String ORvalname = "";
	static String ORvalue = "";
	static String Action = "";
	static String cCellData = "";
	static String dCellData = "";
	static String htmlname = "";
	String[] cCellDataVal = null;
	String[] dCellDataVal = null;
	String ObjectSet;
	String ObjectSetVal = "";
	static Sheet DTsheet = null;
	static Sheet ORsheet;
	String Searchtext;
	static int iflag = 0;
	static int j = 0;
	static int reporttype = 0;
//	static float j = 0;
	int loopsize = -1;
	int[] loopstart = new int[1];
	int[] loopcount = new int[1];
	int[] loopend = new int[1];
	int[] loopcnt = new int[1];
	static int[] dtrownumloop = new int[1];
	boolean captureperform = false;
	static boolean capturecheckvalue = false;
	static boolean capturestorevalue = false;
	static Sheet TScsheet;
	static Workbook TScworkbook;
	static int TScrowcount = 0;
	static int loopnum = 1;
	static String TScname;
	static String ActionVal;
	static String BrowserType; //= "FF"; // Assign with either FF or IE or AD (for Android)
	static String WorkingFolder = "C:\\Shashavali\\DSMB_Workspace\\Shah\\POC\\";
	//static String WorkingFolder = "D:\\S&P_Automation\\RampOnline_Automation\\"; // kreddy
	static String ORDelimiter = "=!";  // kreddy
	static String ORParamDelimiter = "@@@@@";
	static int DTcolumncount = 0;
	static WebElement elem = null;
	static List<WebElement> elems = null;
	private static Map<String, String> map = new HashMap<String, String>();
	private static Map<String, Float> mapint = new HashMap<String, Float>();
	static String OrdNo;

	/*
	 * This function reads the selenium utility file and identifies where Object
	 * Repository, Test Suite & Test Scripts are located
	 */

	@Test
	public void ReadUtilFile() throws Exception 
	{
		{
			PrintWriter pw = new PrintWriter(new FileWriter( WorkingFolder +"test.html"));
			pw.println("<TABLE BORDER><TR><TH>Execute<TH>Keyword<TH>Object<TH>Action<Result></TR>");

			Workbook w1 = null;
			try 
			{
				w1 = Workbook.getWorkbook(new File(WorkingFolder + "SeleniumUtility\\Selenium_Utility.xls"));
			} 
			catch (BiffException e)
			{ // TODO Auto-generated catch block
				e.printStackTrace();
			} 
			catch (IOException e)
			{ // TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			Sheet sheet = w1.getSheet(0);
			TestSuite = sheet.getCell(1, 1).getContents();
			TestScript = sheet.getCell(1, 2).getContents();
			ObjectRepository = sheet.getCell(1, 3).getContents();
			TestSummaryReport = sheet.getCell(1, 4).getContents();
			ReportsPath = sheet.getCell(1, 5).getContents();
			TestReport = sheet.getCell(1, 6).getContents();
			ReusableComponents = sheet.getCell(1, 7).getContents();
			TestData = sheet.getCell(1, 8).getContents();
			BrowserType = sheet.getCell(1, 9).getContents();
		}
		
		for (int z = 0; z < 1; z++) 
		{
			loopstart[z] = 0;
			loopend[z] = 0;
			loopcnt[z] = 0;
			dtrownumloop[z] = 1;
			loopcount[z] = 0;
		}
		switch (BrowserType.toUpperCase())
		{
		case "IE":
			File IDDriver32filePath;
			String OSArchitecture = System.getProperty("os.arch");
			if (OSArchitecture.equalsIgnoreCase("x86"))
			{
				IDDriver32filePath = new File("BrowserDrivers\\IEDriverServer32", "IEDriverServer.exe"); 
			}
			else
			{
				IDDriver32filePath = new File("BrowserDrivers\\IEDriverServer64", "IEDriverServer.exe"); 
			}

			System.setProperty("webdriver.ie.driver", IDDriver32filePath.getAbsolutePath()); 
			DesiredCapabilities capability = DesiredCapabilities.internetExplorer();			
			capability.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,true);
			capability.setCapability("useLegacyInternalServer", true);
			D8 = new InternetExplorerDriver(capability);
			// D8.getWindowHandle();
			D8.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
			D8.manage().window().maximize();
			break;
		case "FF":
			ProfilesIni profile = new ProfilesIni();
			FirefoxProfile ffprofile = profile.getProfile("default");
			D8 = new FirefoxDriver(ffprofile);
			D8.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
			D8.manage().window().maximize();
			break;
		case "AD":
			D8 = new AndroidDriver();
		case "CHROME":
				System.setProperty("webdriver.chrome.driver", WorkingFolder + "\\BrowserDrivers\\chromedriver_win32\\chromedriver.exe"); //System.getProperty("user.dir"+"\\BrowserDrivers\\chromedriver_win32") + "\\chromedriver.exe");
				D8 = new ChromeDriver();
				D8.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
				D8.manage().window().maximize();
								
		}
		
		{
		
			FindExecTestscript(TestSuite, TestScript, ObjectRepository);
		}
		
	}

	public void FindExecTestscript(String TestSuite, String TestScript,
			String ObjectRepository) throws Exception 
			{
				System.out.println("Executed");
				try 
					{
					int TSrowcount = 0;
					FileInputStream fs = null;
					WorkbookSettings ws = null;
					fs = new FileInputStream(new File(TestSuite));
					ws = new WorkbookSettings();
					ws.setLocale(new Locale("en", "EN"));
					Workbook TSworkbook = Workbook.getWorkbook(fs, ws);
					Sheet TSsheet = TSworkbook.getSheet(0);
					TSrowcount = TSsheet.getRows();
					cur_dt = new Date();
					DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
					String strTimeStamp = dateFormat.format(cur_dt);
					String rp = ReportsPath;
					if (rp == "") { // if results path is passed as null, by
						// default 'C:\' drive is used
						rp = "C:\\";
					}
		
					if (rp.endsWith("\\")) { // checks whether the path ends with
						// '/'
						rp = rp + "\\";
					}
					// TCNm = scriptName.split("\\.");
					strResultPath = rp + "Log" + "/";
					
					if (TestSummaryReport == "")
					{
						TestSummaryReport = "C:\\";
					}
						
					if (TestSummaryReport.endsWith("\\"))
					{
						TestSummaryReport = TestSummaryReport + "\\";
					}
					
		//			String htmlname1 = rp + "Log" + "/Test_Suite_" + strTimeStamp + ".html";
					//String htmlname1 = TestSummaryReport + "Test_Suite_" + strTimeStamp + ".html";

					String htmlname1 = TestSummaryReport + strTimeStamp + ".html";
					File f = new File(strResultPath);
					f.mkdirs();
					
					bw1 = new BufferedWriter(new FileWriter(htmlname1));
					bw1.write("<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
					bw1.write("<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
					
					bw1.write("<TR><TD BGCOLOR=#FFFFFF WIDTH=15%><img src=C:\\Shashavali\\DSMB_Workspace\\Shah\\LogoTelefonica.jpg alt=Telefonica style=width:200px;height:75spx></TD>"
							+ "<TD COLSPAN=6 BGCOLOR=#FFFFFF><FONT FACE=VERDANA COLOR=BLUE SIZE=3><I>Automation Test Summary Report</I></FONT></TD>"
							+ "<TD COLSPAN=6 BGCOLOR=#FFFFFF><FONT FACE=VERDANA COLOR=BLUE SIZE=3><I>Execution Date & Time :" + strTimeStamp +" </I></FONT></TD>"
							+ "</TR></TABLE>");
					bw1.write("<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%><P></P><P></P><P></P><P></P><P></P></TABLE>");
					bw1.write("<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
					bw1.write("<TR><TD COLSPAN=6 BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Testcase Name</B></FONT></TD>"
							+ "<TD BGCOLOR=#66699 WIDTH=27%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Status</B></FONT></TD>"
							+ "<TD BGCOLOR=#66699 WIDTH=27%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Detail Report</B></FONT></TR>");
					for (int i = 0; i < TSsheet.getRows(); i++) {
						String TSvalidate = "r";
						if (((TSsheet.getCell(0, i).getContents())
								.equalsIgnoreCase(TSvalidate) == true)) {
							// String TCStatus = "Pass";
							String ScriptName = TSsheet.getCell(1, i).getContents();
							
							ExecKeywordScript(ScriptName, TestScript, ObjectRepository);
							String url = DR_URL;
							
							if (exeStatus.equalsIgnoreCase("Failed"))
							{
								bw1.write("<TR><TD COLSPAN=6 BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>"
										+ TCNm[0]
										+ "</B></FONT></TD><TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=RED SIZE=2><B>"
										+ exeStatus + "</B></FONT></TD>"
										+"<TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=GREEN SIZE=2><A href="+url+">Click to View Detail Report</A></TD</TR>");
								TC_FAIL = TC_FAIL + 1;
							} 
							else 
							{
								bw1.write("<TR><TD COLSPAN=6 BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>"
										+ TCNm[0]
										+ "</B></FONT></TD><TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>"
										+ exeStatus + "</B></FONT></TD>"
										+ "<TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=GREEN SIZE=2><A href="+url+">Click to View Detail Report</A></TD></TR>");
								TC_PASS = TC_PASS + 1;
							}
						}
						/*
						 * else { System.out.println(TSvalidate); }
						 */
					}
				
					
					//bw1.write("<FOOTER BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
					bw1.write("<TR><TD COLSPAN=6 BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>Total No. Of TC Executed</B></FONT></TD>"
							+ "<TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>Total No. Of TC PASSED :- "+TC_PASS+"</B></FONT></TD>"
							+ "<TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=RED SIZE=2><B>Total No. Of TC FAILED :- "+TC_FAIL+"</B></FONT></TR></FOOTER>");
					
					bw1.close();
				} 
			catch (Exception e)
			{
			//bw.close();
			System.out.println(e);
			bw1.close();
			}
	}

	public void ExecKeywordScript(String scriptName, String TestScript,
			String ObjectRepository) throws Exception {

		// Report header
		cur_dt = new Date();
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		String strTimeStamp = dateFormat.format(cur_dt);

		if (ReportsPath == "") { // if results path is passed as null, by
									// default 'C:\' drive is used
			ReportsPath = "C:\\";
		}

		if (ReportsPath.endsWith("\\")) { // checks whether the path ends with
											// '/'
			ReportsPath = ReportsPath + "\\";
		}
		TCNm = scriptName.split("\\.");
		strResultPath = ReportsPath + "Log" + "/" + TCNm[0] + "/";
		String htmlname = ReportsPath + "Log" + "/" + TCNm[0] + "/"
				+ strTimeStamp + ".html";
		DR_URL = htmlname;
		
		File f = new File(strResultPath);
		f.mkdirs();
		bw = new BufferedWriter(new FileWriter(htmlname));
		bw.write("<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TR><TD BGCOLOR=#FFFFFF><img src=D:\\DSMB_Workspace\\Shah\\LogoTelefonica.jpg alt=Telefonica style=width:200px;height:75spx></TD><TD BGCOLOR=#FFFFFF 	</TD></TR>");
		bw.write("<TR><TD BGCOLOR=#66699 WIDTH=10%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B></B></FONT></TD><TD BGCOLOR=#66699 WIDTH=27%></TD></TR>");
		bw.write("<TR><TD BGCOLOR=#66699 WIDTH=10%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Test Case Name:</B></FONT></TD><TD COLSPAN=6 BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>"
				+ TCNm[0] + "</B></FONT></TD></TR>");
		bw.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TR COLS=6><TD BGCOLOR=#FFCC99 WIDTH=3%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Row</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99 WIDTH=15%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Keyword</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Object</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Action</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Execution Time</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Status</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Snapshot</B></FONT></TD></TR>");

		exeStatus = "Pass";
		String scriptPath = TestScript + scriptName;
		TScname = scriptName;
		FileInputStream fs1 = null;
		WorkbookSettings ws1 = null;
		fs1 = new FileInputStream(new File(scriptPath));
		ws1 = new WorkbookSettings();
		ws1.setLocale(new Locale("en", "EN"));
		Workbook TScworkbook = Workbook.getWorkbook(fs1, ws1);
	//	Sheet TScsheet = TScworkbook.getSheet(0);
		TScsheet = TScworkbook.getSheet(0);
		TScrowcount = TScsheet.getRows();
		// *This is the Data Table Sheet
		rowcnt = 0;
		// System.out.println("Row count : " + TScrowcount);
		for (j = 0; j < TScrowcount; j++) {
			// Thread.sleep(1000);
			// System.out.println("J : " + j);
			rowcnt = rowcnt + 1;
			String TSvalidate = "r";
			if (((TScsheet.getCell(0, j).getContents())
					.equalsIgnoreCase(TSvalidate) == true)) {
				Action = TScsheet.getCell(1, j).getContents();
				
				cCellData = TScsheet.getCell(2, j).getContents();
				dCellData = TScsheet.getCell(3, j).getContents();
				String ORPath = ObjectRepository;
				FileInputStream fs2 = null;
				WorkbookSettings ws2 = null;
				try {
					fs2 = new FileInputStream(new File(ORPath));
					ws2 = new WorkbookSettings();
					ws2.setLocale(new Locale("en", "EN"));
				} catch (Exception e) {
					System.out.println("File not found");
				}
				try {
					Workbook ORworkbook = Workbook.getWorkbook(fs2, ws2);
					ORsheet = ORworkbook.getSheet(0);
					ORrowcount = ORsheet.getRows();
					ActionVal = Action.toLowerCase();
					iflag = 0;

				} catch (Exception e) {
					// System.out.println(e);
					fail("Excel file of Open2Test is not correct.");
				}
				System.out.println(Action +"||"+ cCellData +"||"+ dCellData );

                 bcellAction(scriptName);
//				waitForLoad (D8);
			}// End of Execution

		}// End of If that get all rows in Test Script
		bw.close();
	}// End of For that get all rows in Test Script

	public static void screenshot(int loopn, int rown, String Sname) throws Exception
	
	{
		
		String tcScfoldername = null;
				
		try {
			
			//______________________________
			//__________________________________
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
			Date date = new Date();
			String strTime = dateFormat.format(date);
			Sname = Sname.substring(0, Sname.indexOf("."));
			TestReport = TestReport.toLowerCase();
			if (TestReport == "")
				TestReport = ReportsPath;
			if (!(TestReport.contains("screen")))
				TestReport = TestReport + "Screenshot/";
			//-------------------------------------------------- KREDDY
				tcScfoldername = TestReport + Sname ;
				
				File f = new File(tcScfoldername);
				
				if (f.exists() == false)
				{
					f.mkdirs();
				}
		   //--------------------------------------------------	KREDDY
			filenamer = TestReport + Sname + "/" + Sname + "_rowno_"
					+ (j + 1) + "_" + strTime + ".png";			
			Thread.sleep(1000);
			BufferedImage image = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
		    ImageIO.write(image, "png", new File(filenamer));
//		    FileUtils.copyFile(screenshot, new File(filenamer));
		    
		   //__________________________________
//			FileUtils.copyFile(screenshot, new File(filenamer));
		} catch (Exception e) {
			System.out.println(e);
			// System.out.println("Getting Screenshot is failed. Please confirm the test report whether the operation is executed or not.");
			// System.out.println("This message may be displayed when closing the dialog.");
		}
	}
	

	public static void Update_Report(String Res_type) throws IOException {
		String str_time;
		String[] str_rep = new String[2];
		Date exec_time = new Date();
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		str_time = dateFormat.format(exec_time);
		String Allign;
		if (reporttype == 1)
		{
			 Allign =  "middle";
		}else
		{
			 Allign =  "left";
		}
		
		if (Res_type.startsWith("executed")) {
			bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5% ALIGN='"+Allign+"'><FONT FACE=VERDANA SIZE=2>"
					+ (j )
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
					+ Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ cCellData
					+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ dCellData
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = GREEN>"
					+ "Passed" + "</FONT></TD></TR>");						
		} else if (Res_type.startsWith("failed")) {
			exeStatus = "Failed";
			bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5% ALIGN='"+Allign+"'><FONT FACE=VERDANA SIZE=2>"
					+ (j )
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
					+ Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ cCellData
					+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ dCellData
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
					+ "Failed" + "</FONT></TD>"
					+ " </FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
					+ "Snapshot" + "</FONT></TD></TR>");
		} else if (Res_type.contains("CallComponent")) // Kreddy
		{
			bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=BLUE><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = BLUE>"
					+ Res_type + "</div></th></FONT></TR>");		
		
		}
		else if (Res_type.startsWith("loop")) {
			bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=BLUE><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = BLUE>"
					+ Res_type + "</div></th></FONT></TR>");
		} else if (Res_type.startsWith("missing")) {
			exeStatus = "Failed";
			bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5% ALIGN='"+Allign+"'><FONT FACE=VERDANA SIZE=2>"
					+ (j )
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
					+ Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ cCellData
					+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ dCellData
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
					+ "Failed" + "</FONT></TD></TR>");
			bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
					+ (j)
					+ ".Description: The Datatable column name not found</div></th></FONT></TR>");
		} else if (Res_type.startsWith("ObjectLocator")) {
			bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5% ALIGN='"+Allign+"'><FONT FACE=VERDANA SIZE=2>"
					+ (j )
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
					+ Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ cCellData
					+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ dCellData
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
					+ "Failed" + "</FONT></TD></TR>");
			bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
					+ (j )
					+ ".Description: Object Locator is wrong or not supported. Supported Locators are Id,Name,Xpath& CSS</div></th></FONT></TR>");
			}
		
	}
	
	public static void Update_Report(String Res_type, Exception e) throws Exception
	
	{
		screenshot(loopnum, TScrowcount, TScname); // Added by shashavali
		System.out.println("Failed Screenshot  " + filenamer );
		String str_time;
		Date exec_time = new Date();
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		str_time = dateFormat.format(exec_time);
		exeStatus = "Failed";
		String Allign ;
		if (reporttype == 1)
		{
			 Allign =  "middle";
		}else
		{
			 Allign =  "left";
		}
		if (Res_type.startsWith("failed")) {
			bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5% ALIGN='"+Allign+"'><FONT FACE=VERDANA SIZE=2>"
					+ (j )
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
					+ Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ cCellData
					+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ dCellData
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
					+ "Failed" + "</FONT></TD>"
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
					+ "<A href= "+filenamer+">Snapshot</A>" + "</FONT></TD></TR>");
			bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
					+ e.toString().substring(
							e.toString().indexOf(":") + 1,
							e.toString().indexOf(".",
									e.toString().indexOf(":") + 1) + 1)
					+ "</div></th></FONT></TR>");
		} 
	}

	@SuppressWarnings("null")
	private static void Func_StoreCheck() throws Exception {
		// TODO Auto-generated method stub
		try {
			String actval = null;
			String expval;
			Boolean boolval = null;
			String varname;
			String[] cCellDataValCh = cCellData.split(";");
			String ObjectValCh = cCellDataValCh[1];
			String[] dCellDataValCh = dCellData.split(":");
			String ObjectSetCh = dCellDataValCh[0];
			String ObjectSetValCh = "";
			int DTcolumncountCh = 0;
			// DTcolumncountCh = DTsheet.getColumns();
			if (dCellDataValCh.length == 2) {
				ObjectSetValCh = dCellDataValCh[1];
			}
			

			if (ObjectValCh.contains("#")) // Added by Kreddy
			{
				String[] ORcellData = map.get(ObjectValCh.substring(1, (ObjectValCh.length()))).split(ORDelimiter);
				ORvalname = ORcellData[0];
				ORvalue = ORcellData[1];
			}else // Moved below logic to Else case by Kreddy 
			{
				for (int k = 0; k < ORrowcount; k++) {
					String ORName = ORsheet.getCell(1, k).getContents();

					if (((ORsheet.getCell(1, k).getContents())
							.equalsIgnoreCase(ObjectValCh) == true)) {
						String[] ORcellData = new String[3];
						ORcellData = (ORsheet.getCell(4, k).getContents())
								.split(ORDelimiter);
						ORvalname = ORcellData[0];
						ORvalue = ORcellData[1];
						k = ORrowcount + 1;
					}
				}
			}
			
			
			
			if (ObjectSetValCh.contains("dt_")) {
				String ObjectSetValtableheader[] = ObjectSetValCh.split("_");
				int column = 0;
				String Searchtext = ObjectSetValtableheader[1];

				for (column = 0; column < DTsheet.getColumns(); column++) {  //Replaced DTcolumncount with DTsheet.getColumns()  by Kreddy
					if (Searchtext.equalsIgnoreCase(DTsheet.getCell(column, 0)
							.getContents()) == true) {
						ObjectSetValCh = DTsheet.getCell(column, dtrownum)
								.getContents();
						iflag = 1;
					}
				}
				if (iflag == 0) {
					ORvalname = "exit";
				}
			}
			switch (ObjectSetCh.toLowerCase()) {
			case "enabled":
				Func_FindObj(ORvalname, ORvalue);
				boolval = elem.isEnabled();
				actval = boolval.toString();
				break;
				
			case "text":
				Func_FindObj(ORvalname, ORvalue);
				//actval = elem.getAttribute("value");
				actval = elem.getText();
				OrdNo = actval;
				//actval = elem.getAttribute("text");
			
				break;
				
			case "CaptureOrdNumber": // Added by Shashavali
				Func_FindObj(ORvalname, ORvalue);
				actval = elem.getText();
				//System.out.println("Order Number is :- " + actval );
				break ;
				
			
			case "value":
				Func_FindObj(ORvalname, ORvalue);
				actval = new Select(elem).getFirstSelectedOption().getText()
						.toString();
				break;
				
			case "visible":
		
				Func_FindObj(ORvalname, ORvalue);	
				boolval = elem.isDisplayed();
				actval = boolval.toString();
				
				break;
			
						
			//case "invisible":
					// Func_FindObjNew(ORvalname, ORvalue);						
					//boolval = elem.isDisplayed();
					//actval = boolval.toString();
				
				//boolval = elem.isDisplayed();
				//actval = boolval.toString();
				
				//break;
				
				
			case "checked":
				Func_FindObj(ORvalname, ORvalue);
				boolval = elem.isSelected();
				actval = boolval.toString();
				break;
			case "griditems":
				String griditemValue = null;
				Func_FindObjs(ORvalname, ORvalue);
				

				for (WebElement element: elems)
				{
					griditemValue =element.getText(); 
					//System.out.println(griditemValue);
					if (!griditemValue.equalsIgnoreCase(ObjectSetValCh))
					{
						Update_Report("failed");
					}					
				}
				return;	
			case "listitems":
				String listItemvalues = null;
				int intloop = 0;
				//Func_FindObj(ORvalname, ORvalue);
				//List<WebElement> allElements = elem.findElements(By.tagName("li"));
				
				Func_FindObjs(ORvalname, ORvalue);
				//List<WebElement> allElements = elem.findElements(By.tagName("li"));
				

				for (WebElement element: elems)
				{
					if (intloop == 0)
					{
						listItemvalues = element.getText();
						intloop = intloop+1;
					}else
					{
						listItemvalues = listItemvalues + ";" + element.getText();
						intloop = intloop+1;
					}
				}
				actval = listItemvalues.replace(" ", "");
				//actval = listItemvalues;
				break;
			case "linktext":
				Func_FindObj(ORvalname, ORvalue);
				actval = elem.getText();
				break;
			case "elementtext":
				Func_FindObj(ORvalname, ORvalue);
				actval = elem.getText();
				break;
			case "exist":
				Func_FindObj(ORvalname, ORvalue);
				Update_Report("executed");
				return;
			default:
				actval = "Invalid syntax";
				break;
			}
			
			

			if ((ActionVal).equalsIgnoreCase("check")) {
				expval = ObjectSetValCh;
				if (expval.equalsIgnoreCase("On"))
					expval = "True";
				else if (expval.equalsIgnoreCase("Off"))
					expval = "False";
				if (expval.equalsIgnoreCase(actval)) {
					System.out
							.println("Actual value matches with expected value. Actual value is "
									+ actval);
					Update_Report("executed");
				} else {
					System.out
							.println("Actual value doesn't match with expected value. Actual value is "
									+ actval);
					if (ORvalname == "exit") {
						Update_Report("missing");
					} else {
						Update_Report("failed");
					}
					if (capturecheckvalue == true) {
						screenshot(loopnum, TScrowcount, TScname);
						//Mobscreenshot(loopnum, TScrowcount, TScname, D8);
					}
				}
			} else if ((ActionVal).equalsIgnoreCase("storevalue")) {
				System.out.println(actval);
				
				varname = ObjectSetValCh;
				if (actval.equalsIgnoreCase("Invalid syntax")) {
					Update_Report("missing");
				} else {
					if (map.containsKey(varname)) {
						map.put(varname, actval);
						Update_Report("executed");
						System.out
								.println("Overwriting the value of the variable "
										+ varname
										+ " to store the value as mentioned in the test case row number"
										+ rowcnt);
						map.remove(varname);
					} else {
						map.put(varname, actval);
						Update_Report("executed");
						System.out
								.println("Overwriting the value of the variable "
										+ varname
										+ " to store the value as mentioned in the test case row number"
										+ rowcnt);
						if (ORvalname == "exit") {
							Update_Report("missing");
						} else {
						}
					}
				}
				if (capturestorevalue == true) {
					screenshot(loopnum, TScrowcount, TScname);
				}
			}
		} catch (Exception e) {
			 Update_Report("failed", e);
			 bw.close();
		}
	}

	
	
	@AfterTest
	public void close() throws Exception {
		try {
			System.out.println("Test end.");
//			 D8.quit();
		} catch (UnhandledAlertException e) {
			System.out.println(e);
			System.out.println("Because of specification of SeleniumWebDriver, downloading may be failed.");
			System.out.println("Please confirm the report file and screenshot about test result.");
		}
	}

	private static void Func_Authentication_Required() throws Exception
	{
			switch (BrowserType.toUpperCase()) 
			{
				case "IE":					
					try
					{
						File file = new File("Jars\\Jacob", "jacob-1.17-M4-x86.dll"); //path to the jacob dll 
						System.setProperty(LibraryLoader.JACOB_DLL_PATH, file.getAbsolutePath()); 
						File WinSecFile = new File("Jars\\AutoItPopUpHandlers", "popup1.exe");
						String WinSecFilePopupPath = WinSecFile.getAbsolutePath();
						
						AutoItX x = new AutoItX();	
												
						while (true)
						{
							if (x.winExists("Windows Security") == true)
							{
								break;
							}else
							{
								Thread.sleep(2000);
							}
						}
						
						
						if (x.winExists("Windows Security"))
						{
							x.winWaitActive("Windows Security");
							screenshot(loopnum, TScrowcount, TScname);
							Runtime.getRuntime().exec(WinSecFilePopupPath);
							screenshot(loopnum, TScrowcount, TScname);
//							waitForLoad (D8);
//							Thread.sleep(5000);
						}				
						break;
						
					}					
					catch (Exception e)
					{
						try {
							Update_Report("failed", e);
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}	
						break;
					}	
					
				case "FF":
					try
					{
						File file = new File("Jars\\Jacob", "jacob-1.17-M4-x86.dll"); //path to the jacob dll 
						File AuthReqFile = new File("Jars\\AutoItPopUpHandlers", "PopupHandling.exe");
						String AuthReqFilePopupPath = AuthReqFile.getAbsolutePath();
						System.setProperty(LibraryLoader.JACOB_DLL_PATH, file.getAbsolutePath()); 
						
						AutoItX x = new AutoItX();	
						
						if (x.winExists("Authentication Required"))
						{
							x.winActivate("Authentication Required");
							screenshot(loopnum, TScrowcount, TScname);
							Runtime.getRuntime().exec(AuthReqFilePopupPath);
							Thread.sleep(2000);
							screenshot(loopnum, TScrowcount, TScname);
							D8.switchTo().alert().accept();		
							
							Thread.sleep(5000);
						}

					}
					catch (Exception e)
					{
						try
						{
							Update_Report("failed", e);							
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();							
						}							
					}
					break;
			}				
	}
	private static void Func_FindObj(String strObjtype, String strObjpath)
			throws Exception {
		try {
			if (strObjtype.length() > 0 && strObjpath.length() > 0) {
				
				if (strObjtype.equalsIgnoreCase("id")) {
					elem = D8.findElement(By.id(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("name")) {
					elem = D8.findElement(By.name(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("xpath")) {					
					elem = D8.findElement(By.xpath(strObjpath));					
				} else if (strObjtype.equalsIgnoreCase("link")) {
					elem = D8.findElement(By.linkText(strObjpath.toString()));
					//elem = D8.findElement(By.linkText(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("css")) {
					elem = D8.findElement(By.cssSelector(strObjpath));
				} 
			}
		}
		catch (Exception e) {
			//e.printStackTrace();
			Update_Report("failed", e);
			elem = null;
		}
	}
	/*
    private static void Func_FindObjNew(String strObjtype, String strObjpath) throws IOException{

		try {
			if (strObjtype.length() > 0 && strObjpath.length() > 0) {
				if (strObjtype.equalsIgnoreCase("id")) {
					elem = D8.findElement(By.id(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("name")) {
					elem = D8.findElement(By.name(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("link")) {
				elem = D8.findElement(By.linkText(strObjpath.toString()));
				} else if (strObjtype.equalsIgnoreCase("css")) {
					elem = D8.findElement(By.cssSelector(strObjpath));
					
				}
				
			}
		}
		catch (Exception e) {
			e.printStackTrace();
			Update_Report("failed", e);
			// System.out.println(e.toString());
			elem = null;
		}
			
    }*/
	private static void Func_FindObjs(String strObjtype, String strObjpath)throws Exception
			 {
		try {
			if (strObjtype.length() > 0 && strObjpath.length() > 0) {
				if (strObjtype.equalsIgnoreCase("id")) {
					elems = D8.findElements(By.id(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("name")) {
					elems = D8.findElements(By.name(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("xpath")) {
					elems = D8.findElements(By.xpath(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("link")) {
					elems = D8.findElements(By.linkText(strObjpath.toString()));
				} else if (strObjtype.equalsIgnoreCase("css")) {
					elems = D8.findElements(By.cssSelector(strObjpath));
				}
			}
		} catch (Exception e) {
			//Update_Report("failed", e);
			// System.out.println(e.toString());
			elem = null;
		}
	}
	public static int ifContidionSkipper(String strifConditionStatus)
			throws Exception {
		try {

			String strKeyword;
			int intLogicStartRow, intLogicEndRow, intIfEndConditionCount, intIfConditionCount;
			String strKeyWord;
			intIfConditionCount = 1;
			intIfEndConditionCount = 0;
			if (strifConditionStatus.equalsIgnoreCase("false")) {
				intLogicStartRow = j;
				do {
					j = j + 1;
										
					strKeyword = TScsheet.getCell(1, j).getContents();
					System.out.println(strKeyword);
					if (strKeyword.equalsIgnoreCase("Condition")) {
						intIfConditionCount = intIfConditionCount + 1;
					}
					if (strKeyword.equalsIgnoreCase("Endcondition")) {
						intIfEndConditionCount = intIfEndConditionCount + 1;
						if (intIfConditionCount == intIfEndConditionCount) {
							j = j + 1;
							break;
						}
					}

				} while (true);
			}
		} catch (Exception e) {

		}
		return j;
	}

	public String Func_IfCondition(String strConditionArgs) throws Exception {
		int iFlag = 1;
		String str[] = strConditionArgs.split(";");
		String value1 = str[0];
		String value2 = str[2];
		String strOperation = str[1];
		strOperation = strOperation.toLowerCase().trim();

		if (value1.contains("dt_"))
		{
			value1 = GetTestData (value1);
		}
		
		if (value2.contains("dt_"))
		{
			value2 = GetTestData (value2);
		}
		
		if (strOperation.contains("dt_"))
		{
			strOperation = GetTestData (strOperation);
		}
		
		
		
		switch (strOperation.toLowerCase()) {
		case "equals":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				value1 = map.get(value1.substring(1, (value1.length())));
				System.out
						.println("Variable used in condition statement has value: "
								+ value1);
				if (value1.trim().equalsIgnoreCase(value2.trim())) {
					iFlag = 0;
				}
			} else if (value1.trim().equalsIgnoreCase(value2.trim())) {
				iFlag = 0;
			}
			break;

		case "notequals":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				value1 = map.get(value1.substring(1, (value1.length())));
				System.out
						.println("Variable used in condition statement has values: "
								+ value1);
				if (value1.trim().equalsIgnoreCase(value2.trim())) {
					iFlag = 0;
				}
			} else if (!value1.trim().equals(value2.trim())) {
				iFlag = 0;
			}
			break;

		case "greaterthan":
			if (isInteger(value1) && isInteger(value2)) {
				if (Integer.parseInt(value1) > Integer.parseInt(value2)) {
					iFlag = 0;
				}
			} else {
				// Report
			}
			break;
		case "lessthan":
			if (isInteger(value1) && isInteger(value2)) {
				if (Integer.parseInt(value1) < Integer.parseInt(value2)) {
					iFlag = 0;
				}
			} else {
				// Report
			}
			break;
		case "contains":
			if (value1.contains(value2))
			{
				
					iFlag = 0;
				
			} else {
				// Report
			}
			break;	
		case "notcontains":
			if (!value1.contains(value2))
			{
				
					iFlag = 0;
				
			} else {
				// Report
			}
			break;	
		default:
			Update_Report("missing");

		}
		if (iFlag == 0) {
			return "true";
		} else {
			return "false";
		}

	}

	public void arrayresize() {
		if (loopstart.length <= loopsize) {
			loopstart = Arrays.copyOf(loopstart, loopstart.length + 1);
			loopend = Arrays.copyOf(loopend, loopend.length + 1);
			loopcnt = Arrays.copyOf(loopcnt, loopcnt.length + 1);
			dtrownumloop = Arrays.copyOf(dtrownumloop, dtrownumloop.length + 1);
			loopcount = Arrays.copyOf(loopcount, loopcount.length + 1);
		}
	}

	public void bcellAction(String scriptName) throws Exception {
		try {
			switch (ActionVal.toLowerCase()) {

			case "loop":
				startrow = j;
				dtrownum = 1;
				loopsize = loopsize + 1;
				if (loopsize >= 1) {
					arrayresize();
				}
				loopcount[loopsize] = Integer.parseInt(cCellData);
				loopstart[loopsize] = j;
				loopcnt[loopsize] = 0;
				dtrownumloop[loopsize] = dtrownum;
				Update_Report("Start of loop : " + loopsize);
				Update_Report("executed");
				break;
			case "endloop":
				loopcnt[loopsize] = loopcnt[loopsize] + 1;
				loopnum = loopnum + 1;
				if (loopcnt[loopsize] == loopcount[loopsize]) {
					Update_Report("loop" + " End of Loop : " + (loopsize + 1)
							+ " : Loop count : " + loopcnt[loopsize]);
					loopsize = loopsize - 1;
					if (loopsize >= 0)
						dtrownum = dtrownumloop[loopsize];
					else
						dtrownum = 1;
					Update_Report("executed");
				} else {
					j = loopstart[loopsize];
					dtrownum = dtrownum + 1;
					dtrownumloop[loopsize] = dtrownum;
					Update_Report("loop" + " End of Loop : " + (loopsize + 1)
							+ " : Loop count : " + loopcnt[loopsize]);
				}
				rowcnt = 1;
				break;
			case "callaction":
				O2TDriverScript obj2 = new O2TDriverScript();
				String[] mname = cCellData.split(";");
				String method_name = mname[0];
				String method_attribute = mname[1];
				Method m = obj2.getClass().getMethod(method_name, String.class);
				m.invoke(obj2, method_attribute);
				break;
			case "callcomponent":
				reporttype = 1;
				exeStatus = "Pass";
				String ComponentPath = ReusableComponents + cCellData;
				String ComponentName = cCellData.split(".xls")[0];
				FileInputStream ComponentFile1 = null;
				WorkbookSettings ComponentWS1 = null;
				int ComponentRowCount = 0;
				int introwcnt = 0;	
				int introwcntStore = j;	
				
				Update_Report ( j + " - Start of CallComponent : '" + ComponentName + "' execution");				
				ComponentFile1 = new FileInputStream(new File(ComponentPath));
				ComponentWS1 = new WorkbookSettings();
				ComponentWS1.setLocale(new Locale("en", "EN"));
				Workbook ComponentWorkBook = Workbook.getWorkbook(ComponentFile1, ComponentWS1);
				Sheet ComponentSheet = ComponentWorkBook.getSheet(0);
				ComponentRowCount = ComponentSheet.getRows();
				introwcnt = 0;

				for (int jloop = 0; jloop < ComponentRowCount; jloop++) {

					introwcnt = introwcnt + 1;
//					j = j +1;
					
					j = jloop;
					String CTValidate = "r";
					if (((ComponentSheet.getCell(0, jloop).getContents())
							.equalsIgnoreCase(CTValidate) == true)) {
						Action = ComponentSheet.getCell(1, jloop).getContents();
						cCellData = ComponentSheet.getCell(2, jloop).getContents();
						dCellData = ComponentSheet.getCell(3, jloop).getContents();
						String ORPath = ObjectRepository;
						FileInputStream ComponentFile2 = null;
						WorkbookSettings ComponentWS2 = null;
						try {
							ComponentFile2 = new FileInputStream(new File(ORPath));
							ComponentWS2 = new WorkbookSettings();
							ComponentWS2.setLocale(new Locale("en", "EN"));
						} catch (Exception e) {
							System.out.println("File not found");
						}
						try {
							Workbook ORworkbook = Workbook.getWorkbook(ComponentFile2, ComponentWS2);
							ORsheet = ORworkbook.getSheet(0);
							ORrowcount = ORsheet.getRows();
							ActionVal = Action.toLowerCase();
							iflag = 0;

						} catch (Exception e) {
							// System.out.println(e);
							fail("Excel file of Open2Test is not correct.");
						}
						System.out.println(Action +"||"+ cCellData +"||"+ dCellData );
						bcellAction(scriptName);
					}// End of Execution

				}// End of If that get all rows in Test Script
				
				Update_Report ("    End of CallComponent : '" + ComponentName + "' execution");
//				Update_Report("executed");
				j = introwcntStore;
				reporttype = 0;
				break;
			case "buildobject":
				
				String[] cCellDataValuesToBuild = cCellData.split(";");
				String[] dCellDataValuesToBuild = dCellData.split(":");
				String ObjectValuesToBuild = cCellDataValuesToBuild[0];
				String ObjectValuesToStore = cCellDataValuesToBuild[1];
//				String[] OrObjectProp = GetObjPro (ObjectValuesToBuild).split(ORDelimiter);
//				String[] ObjectToBuild = OrObjectProp[1].split(ORParamDelimiter);
				
				for (int iloop = 0 ; iloop < dCellDataValuesToBuild.length;iloop ++)
				{
					
					if (dCellDataValuesToBuild[iloop].contains("dt_") == true)
					{
						dCellDataValuesToBuild[iloop] = GetTestData(dCellDataValuesToBuild[iloop]);
					}
					
				}
				
				if (dCellDataValuesToBuild.length ==1 && dCellDataValuesToBuild[0].contains(":"))
				{
					dCellDataValuesToBuild = dCellDataValuesToBuild[0].split(":");					
				}
				
				String[] ObjectToBuild = GetObjPro (ObjectValuesToBuild).split(ORParamDelimiter);
				String FinalObject = null;
				int objectlen = ObjectToBuild.length;
				objectlen = objectlen-1;
				
				if (objectlen == dCellDataValuesToBuild.length)
				{
					for (int k = 0; k < dCellDataValuesToBuild.length; k++)
						{
							if (k==0)
								{
									FinalObject = ObjectToBuild[k] + dCellDataValuesToBuild[k]+ObjectToBuild[k+1];
								}else
								{
									FinalObject = FinalObject + dCellDataValuesToBuild[k]+ObjectToBuild[k+1];
								}
						}
					
				}
				
//				if (map.containsKey(ObjectValuesToStore)) {
					map.put(ObjectValuesToStore, FinalObject);
					Update_Report("executed");	
					System.out.println(map.get(ObjectValuesToStore));
//					map.get(ObjectSetVal.substring(1,(ObjectSetVal.length())));
//					map.remove(ObjectValuesToStore);
//					System.out.println(map.get(ObjectValuesToStore));
					
//				} 
//				else {
//					map.put(varname, actval);
//					Update_Report("executed");						
//					if (ORvalname == "exit") {
//						Update_Report("missing");
//					} else {
//					}
//				}
			
			break;	
			case "popuphandler":
				switch (cCellData.toLowerCase()) 
				{
					case "authenticationrequired":
					{								
						Func_Authentication_Required();
						Update_Report("executed");
						break;
					}	
					
					case "logoutpopup":
					{
						switch (BrowserType.toUpperCase()) 
						{
							case "FF":	
							{
								
								File file = new File("Jars\\Jacob", "jacob-1.17-M4-x86.dll"); //path to the jacob dll 
								System.setProperty(LibraryLoader.JACOB_DLL_PATH, file.getAbsolutePath()); 
								AutoItX x = new AutoItX();
								
								File WinSecFile = new File("Jars\\AutoItPopUpHandlers", "LogOut.exe");
								String WinSecFilePopupPath = WinSecFile.getAbsolutePath();
								
								if (x.winExists("The page at http://ratingsgateway-qa.mhf.mhc says:"))
								{
									x.winWaitActive("The page at http://ratingsgateway-qa.mhf.mhc says:");	
									Runtime.getRuntime().exec(WinSecFilePopupPath);
									screenshot(loopnum, TScrowcount, TScname);
									Thread.sleep(5000);
								}	
								Update_Report("executed");
								break;
							}
							case "IE":
							{
								File file1 = new File("Jars\\Jacob", "jacob-1.17-M4-x86.dll"); //path to the jacob dll 
								System.setProperty(LibraryLoader.JACOB_DLL_PATH, file1.getAbsolutePath()); 
								AutoItX x1 = new AutoItX();						
								if (x1.winExists("Message from webpage"))
								{
									x1.winWaitActive("Message from webpage");		
									x1.controlClick("Message from webpage","","OK");
								}												

								Update_Report("executed");
								break;
							}										
					}
					break;
				}	
				
				}
				
			case "PdfPopUp":
			
			{
				File file1 = new File("Jars\\Jacob", "jacob-1.17-M4-x86.dll"); //path to the jacob dll 
				System.setProperty(LibraryLoader.JACOB_DLL_PATH, file1.getAbsolutePath()); 
				AutoItX x2 = new AutoItX();						
				if (x2.winExists("File Download"))
				{
					x2.winWaitActive("File Download");		
					x2.controlClick("File Download","","Open");
				}												

				Update_Report("executed");
				break;
			}		
			
				
			case "callfunction":
				
				switch (cCellData.toLowerCase()) 
				{
					case "ckeditor":						
						{								
							String[] dCellDataValues = dCellData.split(":");
							
							switch (dCellDataValues[0].toLowerCase())
							{
								case "set":
								{
									 ((JavascriptExecutor) D8).executeScript("document.body.innerHTML=''");
									 ((JavascriptExecutor) D8).executeScript("document.body.innerHTML='" + dCellDataValues[1] +"'");
										
									 Update_Report("executed");
									 break;
								}
							}										
							break;
							
						}		
					case "save": // Under Development - Kreddy
					{
						try
						{
							Screen screen = new Screen();
//							String screenobjPath = ObjectRepository.replace("ObjectRepository.xls", "ScreenObjects");
//							String SaveButtonPath = screenobjPath + "/SaveIMG.png";
							String SaveButtonPath =  "D:\\S&P_Automation\\RampOnline_Automation\\ObjectRepository\\ScreenObjects\\SaveIMG.png";
																
							Pattern SaveButton = new Pattern(SaveButtonPath);
//							screen.wait(SaveButton, 3);
							screen.click(SaveButton,10);	
						}catch (Exception e)
						{
							System.out.println( "Sikuli Exception :  " + e.getMessage());
						}
						break;
					}
					
					case "componenthierarchyverification":
					{
						componentHierarchyVerification();
						break;
					}
				}
				break;
				

			case "launchapp":	
				
				System.out.println(cCellData);
				D8.get(cCellData);
				screenshot(loopnum, TScrowcount, TScname);
				Update_Report("executed");
				break;
				
			case "selectmobile": // This is specific to DSMB Project
				
				//System.out.println("Case Selected");
				cCellDataVal = cCellData.split(";");
				
			
				String ExpDeviceName = cCellDataVal[0]+"-"+cCellDataVal[1];
				System.out.println(ExpDeviceName);
				selectgivenmobile(ExpDeviceName);
								
				Update_Report("executed");
				if (captureperform == true) {
					screenshot(loopnum, TScrowcount, TScname);
				}
				break;
				
			case "sikuliclick": // This is specific to DSMB Project
				
				System.out.println("Case Selected");
				System.out.println(cCellData);
				
				sikuliclickimage(cCellData);
								
				Update_Report("executed");
				if (captureperform == true) {
					screenshot(loopnum, TScrowcount, TScname);
				}
				break;
			
			case "checkappname":
				
				String Res;					
				String ExpAppName = cCellData;
				
				Res = checkgivenapp(ExpAppName);
				if(Res.equalsIgnoreCase("TRUE"))
				{
					Update_Report("executed");
					if (captureperform == true) 
					{
						screenshot(loopnum, TScrowcount, TScname);
					}
					else
					{
						Update_Report("failed");
						if (captureperform == true) 
						{
							screenshot(loopnum, TScrowcount, TScname);
						}
					}
				}
				
				break;
				
			case "switchframe": // kreddy

				String iFrame = null;
				String[] frm = null;
				if(cCellData.contains(";"))
				{
					frm = cCellData.split(";");
					iFrame = GetObjPro (frm[1]);					
				}
				
				if (iFrame == null)
				{
					try
					{
						if (Integer.parseInt(frm[1]) >= 0)
						{						
							int FrameIndex = Integer.parseInt(frm[1]);
							D8.switchTo().frame(FrameIndex);
							Update_Report("executed");
							Thread.sleep(2000);
							break;
						}
					}catch(Exception e)
					{
						D8.switchTo().frame(frm[1]);
						Update_Report("executed");
						Thread.sleep(2000);
						break;
					}	
				}
				
				if(iFrame.contains(ORDelimiter))
				{
					String[] ORcellData = iFrame.split(ORDelimiter);
					ORvalname = ORcellData[0];
					ORvalue = ORcellData[1];
					Func_FindObj(ORvalname, ORvalue);
					if (elem != null) 
					{
						D8.switchTo().frame(elem);
						Update_Report("executed");
						Thread.sleep(2000);
						break;
					}					
				}
				break;	
			case "switchoutframe": // kreddy
				D8.switchTo().defaultContent();
				Update_Report("executed");
				Thread.sleep(2000);
				break;
			case "wait":
				String StrWaitTime = cCellData + "000";
				long intWaitTime = Long.parseLong(StrWaitTime);				
				Thread.sleep(intWaitTime);
				Update_Report("executed");
				break;

			case "condition":
				String strConditionStatus = Func_IfCondition(cCellData);
				if (strConditionStatus.equalsIgnoreCase("false")) {
					j = ifContidionSkipper(strConditionStatus);
					//j = j + 1;
					System.out.println(j);
				}
				Update_Report("executed");
				break;

			case "endcondition":
				Update_Report("executed");
				break;

			case "screencaptureoption":
				String[] sco = cCellData.split(";");

				for (int s = 0; s < sco.length; s++) {
					if (sco[s].equalsIgnoreCase("perform")) {
						captureperform = true;
					}
					if (sco[s].equalsIgnoreCase("storevalue")) {
						capturestorevalue = true;
					}
					if (sco[s].equalsIgnoreCase("check")) {
						capturecheckvalue = true;
					}

				}
				Update_Report("executed");
				break;
			case "importdata":
				// Runtime rt = Runtime.getRuntime();
				// Process p = rt.exec("D://AutoITScript/FileDown.exe");
				String xcelpath = TestData + cCellData;
				FileInputStream fs3 = null;
				WorkbookSettings ws3 = null;
				fs3 = new FileInputStream(new File(xcelpath));
				ws3 = new WorkbookSettings();
				ws3.setLocale(new Locale("en", "EN"));
				Workbook DTworkbook = Workbook.getWorkbook(fs3, ws3);
				DTsheet = DTworkbook.getSheet(0);
				int DTrowcount = DTsheet.getRows();
				String DTName;
				Update_Report("executed");
				break;
				
			case "screencapture":
				screenshot(loopnum, TScrowcount, TScname);
				Update_Report("executed");
				break;
			case "check":
				Func_StoreCheck();
				break;
			case "storevalue":
				Func_StoreCheck();
				break;
				
				
			case "pageoperations":
				if (cCellData.equalsIgnoreCase("pageback"))
				{
					screenshot(loopnum, TScrowcount, TScname);
					D8.navigate().back();	
					Thread.sleep(2000);
					screenshot(loopnum, TScrowcount, TScname);
					Update_Report("executed");
					break;	
				}
				
				if (cCellData.equalsIgnoreCase("pagerefresh"))
				{
					screenshot(loopnum, TScrowcount, TScname);
					D8.navigate().refresh();
					Thread.sleep(2000);
					screenshot(loopnum, TScrowcount, TScname);
					Update_Report("executed");
					break;	
				}	
				
			case "javascriptexecutor": // kreddy
				
				cCellDataVal = null;
				cCellDataVal = cCellData.split(";");
						
				ObjectSetVal = GetObjPro (cCellDataVal[1]);
				
				JavascriptExecutor jsexecutor = (JavascriptExecutor) D8;	
				jsexecutor.executeScript(ObjectSetVal);
				Thread.sleep(1000);
				
				Update_Report("executed");
				if (captureperform == true)
				{
					screenshot(loopnum, TScrowcount, TScname);
				}
				break;		
			
			
			case "perform": // kreddy
				
				cCellDataVal = null;
				dCellDataVal = null;
				
				if (cCellData.equalsIgnoreCase("Button;MySave"))
				{
				System.out.println("For Debugging");	
				}
				
				cCellDataVal = cCellData.split(";");
				dCellData.toString();
				
				String ObjectVal = cCellData.substring(cCellDataVal[0].length() + 1, cCellData.length());
				
				// To get the Test data
				
				if (dCellData.contains(":")) {
					dCellDataVal = dCellData.split(":");
					ObjectSet = dCellDataVal[0].toLowerCase();
					ObjectSetVal = GetTestData (dCellDataVal[1]);
				} else {
					ObjectSet = dCellData.toString();
					ObjectSetVal = null;  // kreddy
				}
				
				// To get the Runtime object from Map Object
				
				if (ObjectVal.contains("#"))
				{
					String[] ORcellData = map.get(ObjectVal.substring(1, (ObjectVal.length()))).split(ORDelimiter);
					ORvalname = ORcellData[0];
					ORvalue = ORcellData[1];
				}else {
					
					String[] ORcellData = GetObjPro (ObjectVal).split(ORDelimiter);
					ORvalname = ORcellData[0];
					ORvalue = ORcellData[1];
				}
				
				dCellAction();
				// End of Objectset Switch
				// Update_Report("executed");
			}// End of Actval Switch
		} /*
		 * catch (UnhandledAlertException e) {
		 * 
		 * for (String id : D8.getWindowHandles()) {
		 * System.out.println("WindowHandle: " + D8.switchTo().window(id));
		 * 
		 * } System.out.println(e); System.out .println(
		 * "Because of specification of SeleniumWebDriver, downloading may be failed."
		 * ); System.out .println(
		 * "Please confirm the report file and screenshot about test result.");
		 * }
		 */catch (Exception ex) {
			Update_Report("failed", ex);
			System.out.println(ex);
			System.out.println("------Error Information : Open2Test-------");
			System.out.println("Current Script:" + scriptName);
			System.out.println("Current ScriptPath:" + TestScript);
			System.out
					.println("Using ObjectRepositoryPath:" + ObjectRepository);
			System.out.println("Current Keyword:" + Action);
			System.out.println("Current ObjectDetails:" + cCellData);
			System.out.println("Current ObjectDetailsPath:" + ORvalue);
			System.out.println("Current Action:" + dCellData);
			System.out.println("------Error Information : Open2Test-------");
			fail("Cannot test normally by Open2Test.");
			// return;
		}
	}

	private void fail(String string) {
		// TODO Auto-generated method stub
		
	}

	public void readAttributeforPerform() throws Exception {
		try {
			
			if (ObjectSetVal != null && ObjectSetVal.length() > 0) {
				if (ObjectSetVal.substring(0, 1).equalsIgnoreCase("#")) {
					ObjectSetVal = map.get(ObjectSetVal.substring(1,
							(ObjectSetVal.length())));
				} else if (ObjectSetVal.contains("dt_")) {
					String ObjectSetValtableheader[] = ObjectSetVal.split("_");
					int column = 0;
					String Searchtext = ObjectSetValtableheader[1];
					for (column = 0; column < DTcolumncount; column++) {
						if (Searchtext.equalsIgnoreCase(DTsheet.getCell(column,
								0).getContents()) == true) {
							ObjectSetVal = DTsheet.getCell(column, dtrownum)
									.getContents();
							iflag = 1;
						}

					}
					if (iflag == 0) {
						ORvalname = "exit";
					}
				}
			}
		} catch (Exception e) {
			Update_Report("failed", e);
		}

	}

	public void dCellAction() throws Exception {
		try {
			readAttributeforPerform();
			
			Func_FindObj(ORvalname, ORvalue);
			if (elem == null) {
				return;
			} else {
				switch (ObjectSet.toLowerCase()) {
				
				
				case "set":
										
						// readAttributeforPerform();
						// Func_FindObj(ORvalname, ORvalue);
						elem.clear();
						elem.sendKeys(ObjectSetVal);
						Update_Report("executed");
						if (captureperform == true)
						{
							screenshot(loopnum, TScrowcount, TScname);
						}
										
					break;
				case "listselect":

					Func_FindObj(ORvalname, ORvalue);
					String[] listvalues = ObjectSetVal.split(",");
					List<WebElement> listboxitems = elem.findElements(By
							.tagName("option"));
					Select chooseoptn = new Select(elem);
					chooseoptn.deselectAll();
					for (WebElement opt : listboxitems) { // System.out.println(opt.getText());
						for (int i = 0; i < listvalues.length; i++) {
							if (opt.getText().equalsIgnoreCase(listvalues[i])) {
								// System.out.println(listvalues[i]);
								chooseoptn.selectByVisibleText(opt.getText());
							}
						}
					}
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}

					break;

				case "select":
					// readAttributeforPerform();
					// Func_FindObj(ORvalname, ORvalue);
					new Select(elem).selectByVisibleText(ObjectSetVal);
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}

					break;

				case "check":
					// readAttributeforPerform();
					// Func_FindObj(ORvalname, ORvalue);
					if (elem.isSelected()
							&& dCellDataVal[1].equalsIgnoreCase("On")) {

					} else if (elem.isSelected()
							|| dCellDataVal[1].equalsIgnoreCase("On")) {
						elem.click();
					}
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
					break;

				case "click":			
					elem.click();
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
					break;
				case "mousehover": // kreddy	
//					elem.click();
					Actions act = new Actions(D8);
					act.moveToElement(elem).perform();
			        Thread.sleep(1000);			       		        				
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
					break;	
					
				case "jsclick":  // kreddy
					JavascriptExecutor jsclickexecutor = (JavascriptExecutor) D8;
					jsclickexecutor.executeScript("arguments[0].click();", elem);
					Thread.sleep(1000);
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
					break;
				case "jsidclick":  // kreddy					

					JavascriptExecutor jsidclickexecutor = (JavascriptExecutor) D8;
					String elementIdToClick = elem.getAttribute("id");
					
					jsidclickexecutor.executeScript("var webelement = document.getElementById('" + elementIdToClick +"'); webelement.click();");
					Thread.sleep(1000);

					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
					break;				
										
				case "jsidset": // kreddy
														
					 JavascriptExecutor jsidsetexecutor = (JavascriptExecutor) D8;		
					 String elementIdToset = elem.getAttribute("id"); 
					 jsidsetexecutor.executeScript("document.getElementById('" + elementIdToset +"').value = arguments[0];",ObjectSetVal );				 
					 Thread.sleep(1000);
					 Update_Report("executed");
					 if (captureperform == true)
					 {
						screenshot(loopnum, TScrowcount, TScname);
					 }
					 break;
					 
					 
					//elem.sendKeys(ObjectSetVal);
					// fh
				//	CharSequence text = null;
					
				//	text = ObjectSetVal;
					
					 //elem.sendKeys(text);
					 
					// D8.switchTo().window("sdg");
					 
					// D8.switchTo().alert().accept();
					 
				//	((JavascriptExecutor) D8).executeScript("Ext.getCmp('myWorkSearchTxt').onFocus()");
				//	elem.sendKeys(text);
					
					
//					 JavascriptExecutor jsidsetexecutor = (JavascriptExecutor) D8;		
//					 String elementIdToset = elem.getAttribute("id");
//				
//					 //jsidsetexecutor.executeScript("Ext.getCmp('myWorkSearchTxt').value =",text );	
//					 
//					 jsidsetexecutor.executeScript("document.getElementById('" + elementIdToset +"').value = arguments[0];",ObjectSetVal );				 
//					 Thread.sleep(1000);
//					
//					 
//					 Update_Report("executed");
//					 if (captureperform == true)
//					 {
//						screenshot(loopnum, TScrowcount, TScname);
//					 }
//					 break;
								
				
				}
			}
		} catch (Exception e) {
			Update_Report("failed", e);
		}
	}

	public static boolean isInteger(String input) {
		try {
			Integer.parseInt(input);
			return true;
		} catch (Exception e) {
			return false;
		}
	}
	public static String GetObjPro (String ObjName)
	{
		String ORcellData = null;
//		String[] ORcellData = new String[2];
		for (int k = 0; k < ORrowcount; k++)
		{
			if (((ORsheet.getCell(1, k).getContents()).equalsIgnoreCase(ObjName) == true))
			{				
				ORcellData = ORsheet.getCell(4, k).getContents();
				k = ORrowcount + 1;
				return ORcellData;
			}			
		}
		return null;
	}
	
	public static String GetTestData (String dtColumnName)
	{
		String DTCellData = null;	
				
		if (dtColumnName.contains("dt_")) 
		{
			String ObjectSetValtableheader[] = dtColumnName.split("_");
			
			String Searchtext = ObjectSetValtableheader[1];

			for (int column = 0; column < DTsheet.getColumns(); column++)
			{
				if (Searchtext.equalsIgnoreCase(DTsheet.getCell(column, 0).getContents()) == true)
				{
					DTCellData = DTsheet.getCell(column, dtrownum).getContents();
					return DTCellData;
				}
			}
			
		}		
		
		return dtColumnName;
		
	}
	
	public static void waitForLoad(WebDriver driver) {
	    ExpectedCondition<Boolean> pageLoadCondition = new
	        ExpectedCondition<Boolean>() {
	            public Boolean apply(WebDriver driver) {
	                return ((JavascriptExecutor)driver).executeScript("return document.readyState").equals("complete");
	            }
	        };
	    WebDriverWait wait = new WebDriverWait(driver, 180);
	    wait.until(pageLoadCondition);
	    
	    
//	    Boolean readyStateComplete = false;
//	    while (!readyStateComplete) {
//	    	JavascriptExecutor executor = (JavascriptExecutor) D8;
//	        executor.executeScript("window.scrollTo(0, document.body.offsetHeight)");
//	        readyStateComplete = executor.executeScript("return document.readyState").equals("complete");
//	    }
	    
	}

	private static void componentHierarchyVerification() throws Exception
	{
		String ActualTDToVerify = null; 
		String listItemvalues = null;
		String[] testData = null;
		if (dCellData.contains(":"))
		{
			testData = dCellData.split(":");
			if (testData[1].contains("dt_"))
			{
				ActualTDToVerify = GetTestData (testData[1]);				
			}
			else
			{
				ActualTDToVerify = testData[1];
			}			
		}
		

		String[] ObjectToBuild = GetObjPro (testData[0]).split(ORDelimiter);
		Func_FindObj(ObjectToBuild[0], ObjectToBuild[1]);

		
		// elem = D8.findElement(By.xpath("//div[@id='guidanceStack2']/div[contains(@class,'x-panel-body')]"));
		listItemvalues = elem.getText();
		
		listItemvalues = listItemvalues.replace(" ", "");
		ActualTDToVerify = ActualTDToVerify.replace(" ", "");
		
		if (ActualTDToVerify.equalsIgnoreCase(listItemvalues))
		{
			Update_Report("executed");
		}else
		{
			Update_Report("failed");
		
		}
			

		
	}


	
	public static void Mobscreenshot(int loopn, int rown, String Sname, WebDriver driver)
			throws Exception {
		
		String tcScfoldername = null;
				
		try {
			
			//______________________________
			//__________________________________
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
			Date date = new Date();
			String strTime = dateFormat.format(date);
			Sname = Sname.substring(0, Sname.indexOf("."));
			TestReport = TestReport.toLowerCase();
			if (TestReport == "")
				TestReport = ReportsPath;
			if (!(TestReport.contains("screen")))
				TestReport = TestReport + "Screenshot/";
			//-------------------------------------------------- KREDDY
				tcScfoldername = TestReport + Sname ;
				
				File f = new File(tcScfoldername);
				
				if (f.exists() == false)
				{
					f.mkdirs();
				}
		   //--------------------------------------------------	KREDDY
			String filenamer = TestReport + Sname + "/" + Sname + "_rowno_"
					+ (j + 1) + "_" + strTime + ".png";			
			Thread.sleep(1000);
			File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(screenshot, new File(filenamer));
//		    FileUtils.copyFile(screenshot, new File(filenamer));
		    
		   //__________________________________
//			FileUtils.copyFile(screenshot, new File(filenamer));
		} catch (Exception e) {
			System.out.println(e);
			// System.out.println("Getting Screenshot is failed. Please confirm the test report whether the operation is executed or not.");
			// System.out.println("This message may be displayed when closing the dialog.");
		}
	}


public static void selectgivenmobile(String ExpDeviceToSelect) throws InterruptedException, IOException
	
	{
		WebElement result = D8.findElement(By.id("resultList"));
		
		List<WebElement> ResCount = result.findElements(By.tagName("dt"));
		
		System.out.println("Result Count " + ResCount.size());
		int jCount = ResCount.size()/2;
		System.out.println(jCount);
		for(int iCount=1; iCount<=jCount; iCount++ )
		{
			String xPath1 = "//*[@id='resultList']/div["; 
			String xPath2 = "]/div[1]/dl/dt";
			String xPath = xPath1+(iCount)+xPath2;
			
			String deviceName = D8.findElement(By.xpath(xPath)).getText().trim();
				
			String xP1 = "//*[@id='resultList']/div["; 
			String xP2 = "]/div[1]/dl/dd";
			String xP = xP1+(iCount)+xP2;
			
			String deviceCol = D8.findElement(By.xpath(xP)).getText().trim();
			String ActDevice = deviceName+"-"+deviceCol;
			System.out.println(ActDevice);
			if(ExpDeviceToSelect.equalsIgnoreCase(ActDevice))
			{
				
				String xPa1 = "//*[@id='resultList']/div["; 
				String xPa2 = "]";
				String xPa = xPa1+(iCount)+xPa2;
				WebElement SelectBtn = D8.findElement(By.xpath(xPa));
				SelectBtn.findElement(By.tagName("a")).click();
				System.out.println("Given device selected");
			
				
				break;
			}
			
		}
	}

public static void sikuliclickimage(String ImagePath)
{
	ScreenRegion s = new DesktopScreenRegion();
	Target target = new ImageTarget(new File(ImagePath));
	
	ScreenRegion r = s.wait(target, 8000);
	r = s.find(target);
	
	Canvas canvas = new DesktopCanvas();
	canvas.addBox(r);
	canvas.addLabel(r, "Found it");
	canvas.display(3);
	
	
	if(r == null)
	{
		System.out.println("Not Found");
	}
	else
	{
				
		DesktopMouse mouse = new DesktopMouse();
		mouse.click(r.getCenter());
	}
		
		
	}

public static String checkgivenapp(String ExpAppName)
{
	String Result = "False";
	
	for(int j=1;j<=3;j++)
	{
	String xPath1 = "//div[starts-with(@id,'digtalProductDetailsPanel')]/div[";
	String xPath2 = "]/h3";
	String xPath = xPath1+(j)+xPath2;
	
	List<WebElement> AppNames = D8.findElements(By.xpath(xPath));
	for(int i=0; i<AppNames.size();i++)
	{
		WebElement App = AppNames.get(i);
//		System.out.println(App.getText());
		String Act_App_Name = App.getText();
		if(ExpAppName.equalsIgnoreCase(Act_App_Name))
		{
			Result = "True";
			break;
		}
		
	}
	}
	return Result;
	}

}

