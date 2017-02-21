/*
 * Lớp thực hiện các chức năng autotest với đơn PC
 * @Author: hieuht
 * @Date: 01/12/2016
 */
package DonPC;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.ResultSet;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.Enumeration;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import javax.swing.JOptionPane;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.openqa.jetty.http.nio.SocketChannelOutputStream;
import org.openqa.selenium.By;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebElement;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;
import org.testng.collections.Lists;

import Common.Common;
import Common.DatabaseAction;
import Common.ExcelAction;
import Common.ObjectInput;
import Common.WebDriverAction;
import DonPC.CheckDatabase_PC;

public class FunctionTest_PC {
	/*
	 * driver: đối tượng web
	 * inputFile: đường dẫn file input thao tác
	 * fileResult: đường dẫn file output thao tác
	 * pathFileResult: đường dẫn mở file output
	 * timeWaitAjax: thời gian đợi
	 * db: đối tượng kiểm tra db với data excel
	 * resetData: lựa chọn reset data
	 * dbNameDelete: tên db xóa khi rs data
	 * dbNameImport: tên db import khi rs data
	 */
	private WebDriverAction driver;
	private String File = "DataTest\\DonPC\\Data-DonPC.xls";
	public CheckDatabase_PC db = new CheckDatabase_PC();
	public int resetData = -1;
	private String dbNameDelete = "";
	private String dbNameImport = "";
	private String linkStart = "";
	private String linkLogin;
	private String environment = "";
	private String[] flow;
	private int timeWaitJs = 200;
	private int timeWaitAjax = 0;
	private int noWait = 0;
	private long startTime;
	private long endTime;
	private String buildId;
	private String timeBuild;
	Logger logger = null;
	
	/*
	 * Hàm thực hiện mở trình duyệt FireFox version 46
	 * @Author: hieuht
	 * @Date:01/12/2016
	 */
	public void openBrowser() throws Exception {
		DateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH.mm");
		Date frmDate = sdf.parse(timeBuild); // Handle the ParseException here

		DateFormat sdff = new SimpleDateFormat("yyyy-MM-dd HH.mm");
	    System.setProperty("current.date", sdff.format(frmDate));
	    System.setProperty("current.proposal", "DonPC");
	    System.setProperty("current.build", "#"+buildId);
		//Logger
		logger=Logger.getLogger("openBrower");
		//Config log4j
		PropertyConfigurator.configure("log4j.properties");
		
		Common.readExcelDefine();
		
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		Date date = new Date();
		logger.info("Start autotest don PC at " + dateFormat.format(date));
		startTime = System.currentTimeMillis();
		try{
		//Truy cập vào file input excel sheet 0, đọc các dữ liệu config
		ExcelAction excel = new ExcelAction();
		excel.createFileOutput(File, buildId, timeBuild);
		excel.accessSheet("Config");
		
		int column = 1;		
//		linkStart = excel.getStringData(column, 6);
//		timeWaitAjax = Integer.parseInt(excel.getStringData(column, 9));
		resetData = Integer.parseInt(excel.getStringData(column, 18));
		dbNameImport = excel.getStringData(column, 19);
		dbNameDelete = excel.getStringData(column, 20);
		linkLogin = excel.getStringData(column, 2);
//		environment = excel.getStringData(1, 7);
		
		excel.accessSheet("Config");
		DatabaseAction.setInputFile(File);
		
	    //Mở trình duyệt Firefox
	    driver = new WebDriverAction();
	    driver.getDriver().manage().window().maximize();
	    
	    excel.finish();
	    }catch(FileNotFoundException e){
	    	driver.close();
	    	logger.info("Please close file excel");
	    	JOptionPane.showMessageDialog(null, "Chưa đóng file excel");
	    	System.exit(0);
	    }
	}
	
	/*
	 * Author: hieuht
	 * Đọc tham số truyền vào trên jenkin
	 * Date: 04/01/2017
	 */
		@Test(dataProvider="getData")
		public void setData(String username, String password)
		{
			//call dataprovider
		}
	
		@DataProvider
		public Object[][] getData()
		{
			Object[][] data = new Object[3][2];
			
			String inputEnv = System.getProperty("environment");
			environment = inputEnv;
			data[0][0] = inputEnv;
			if(environment.equals("Test"))
				linkStart = "https://eapply-test.adways.net/";
			else if(environment.equals("Staging"))
				linkStart = "https://eapply-staging.adways.net/";
			
			String inputTime = System.getProperty("timeSleep");
			timeWaitAjax = Integer.parseInt(inputTime);
			data[1][0] = inputTime;
			
			String flowRun = System.getProperty("flowRun");
			flow = flowRun.split(",", -1);
			buildId = System.getProperty("buildID");
			timeBuild = System.getProperty("timeBuild");
			return data;
		}
	
	@Test
	/*
	 * Start
	 * @Author: hieuht
	 * @Date: 01/12/2016
	 */
	public void start() throws Exception{
		try{
		openBrowser();
		//Logger
		logger=Logger.getLogger("start");
		
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);  
		excel.accessSheet("Create Proposal");
		
		excel.accessSheet("Flow");
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "testcase", 0),
				new ObjectInput("", 0, 0, 0, "createProposal", 0),
				new ObjectInput("", 0, 0, 0, "settingApprove", 0),
				new ObjectInput("", 0, 0, 0, "settingBack", 0),
				new ObjectInput("", 0, 0, 0, "approve", 0),
				new ObjectInput("", 0, 0, 0, "back", 0),
				new ObjectInput("", 0, 0, 0, "cancel", 0),
				new ObjectInput("", 0, 0, 0, "statusManagement", 0),
				new ObjectInput("", 0, 0, 0, "createAgain", 0),
				new ObjectInput("", 0, 0, 0, "result", 0),
				new ObjectInput("", 0, 0, 0, "type", 0)};
		try {
			DatabaseAction.doSshTunnel(environment, 0, buildId, timeBuild);	
		} catch (Exception e) {
			logger.error(e.getMessage());
			System.exit(0);
		}
		int resultColumn = Common.getIndexFromArray(array, "result");		
		for(int i=0; i< flow.length; i++){	
			try{
				logger.info("Start test case "+excel.getStringData(Common.getIndexFromArray(array, "testcase"), Integer.parseInt(flow[i])));
				if(doFlow(excel, array, Integer.parseInt(flow[i])) == true)
					excel.printResultToFlow("T", excel.getStringData(Common.getIndexFromArray(array, "type"), Integer.parseInt(flow[i])), excel, Integer.parseInt(flow[i]), resultColumn);
				else{
					excel.printResultToFlow("F", excel.getStringData(Common.getIndexFromArray(array, "type"), Integer.parseInt(flow[i])), excel, Integer.parseInt(flow[i]), resultColumn);
				}
				logger.info("End test case "+excel.getStringData(Common.getIndexFromArray(array, "testcase"), Integer.parseInt(flow[i])));
			}catch(Exception e){
				excel.accessSheet("Flow");
				excel.printResultToFlow("F", excel.getStringData(Common.getIndexFromArray(array, "type"), Integer.parseInt(flow[i])), excel, Integer.parseInt(flow[i]), resultColumn);
				if(driver.getCurrentUrl().equals("https://eapply-test.adways.net/admin/errors/index?error_code=error_action")
						|| driver.getCurrentUrl().equals("https://eapply-staging.adways.net/admin/errors/index?error_code=error_action")) {
					logger.error("Khong lay duoc content don theo ID");
				} else {
					boolean flag = true;
					try {
						driver.getText("className=username_employee_code").equals("");
					} catch (Exception e1) {
						logger.error("Chua dang nhap thanh cong");
						flag= false;
					}
					if(flag) {
						boolean flag1 = false;						
						Set<String> keys = Common.listException.keySet();
						for(String key: keys){
							 if (e.getMessage().startsWith(key)) {
					            	logger.error("#"+ Common.listException.get(key) + "-" + e.getMessage());
					            	flag1 = true;
					            }
					        }
						if(!flag1)
							logger.error(e.getMessage());
					}
				}
				logger.info("End test case "+excel.getStringData(Common.getIndexFromArray(array, "testcase"), Integer.parseInt(flow[i])));
			}
		}
		excel.finish();
		}catch (Exception e) {
			logger.error(e.getMessage());
		}
	}
	
	/*
	 * Gọi đến các chức năng
	 * @Author: hieuht
	 * @Date: 01/12/2016
	 * @Parametter: 
	 * 		ExcelAction excel: file excel
	 * 		ObjectInput[] array: 
	 * 		int row: cột thực hiện auto
	 */
	public boolean doFlow(ExcelAction excel, ObjectInput[] array, int row)throws Exception{
		for(int i=1; i<array.length; i++){
			if(!excel.getStringData(Common.getIndexFromArray(array, array[i].getName()), row).equals("")){
			DatabaseAction.doSshTunnel(environment, 1, buildId, timeBuild);
			int column = excel.getColumn(excel.getStringData(Common.getIndexFromArray(array, array[i].getName()), row));
			switch(array[i].getName()){
				case "createProposal":{
					if(createProposal(column, excel) == false){
						excel.accessSheet("Flow");
						return false;
					}
					else{
						excel.accessSheet("Flow");
						break;
					}
				}
				case "settingApprove":{
					if(settingApprove(column) == false)
						return false;
					else
						break;
				}
				case "settingBack":{
					if(settingBack(column) == false)
						return false;
					else
						break;
				}
				case "approve":{
					if(approve(column) == false)
						return false;
					else 
						break;
				}
				case "back":{
					if(back(column) == false)
						return false;
					else 
						break;
				}
				case "cancel":{
					if(cancel(column) == false)
						return false;
					else 
						break;
				}
				case "statusManagement":{
					if(statusManagement(column) == false)
						return false;
					else
						break;
				}
				case "createAgain":{
					if(createProposalAgain(column) == false)
						return false;
					else
						break;
				}
			}
			excel.write();
			}
		}
		return true;
	}

	@Test
	/*
	 * Hàm thực hiện auto tạo đơn PC
	 * @Author: hieuht
	 * @Date: 01/12/2016
	 * @Parametter: 
	 * 		int column: cột chứa data
	 */
	public boolean createProposal(int column, ExcelAction excel) throws Exception {
		logger=Logger.getLogger("createMultiProposal");

		//Thực hiện reset data nếu dữ liệu trong excel = 2 (đối với mt local)
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
			logger.info("Reset db success");
		}
		excel.accessSheet("Config");
		String linkCreate = excel.getStringData(1, 28);
		excel.accessSheet("Create Proposal");
		
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//1: setvisible, 2: sendkey, 3 pickdate, 4 click, 5 chon setvisible theo chi so, 6 click theo chi so
		//ID, action, type, index, name
		ObjectInput[] array = {	
				new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
				new ObjectInput("", 0, 0, 0, "numberUser", 0),
				new ObjectInput("", 0, 0, 0, "newID", 0),
				new ObjectInput("id=projectdate", 2, 3, 0, timeWaitJs),
				new ObjectInput("id=checkbox_create_by", 2, 7, 0, "helpBy", timeWaitAjax*2),
				new ObjectInput("id=proposal_change", 2, 1, 0, "userHelpBy", timeWaitAjax),
				new ObjectInput("id=company_code", 2, 1, 0, timeWaitAjax),
				new ObjectInput("id=group_code", 2, 1, 0, timeWaitAjax),
				new ObjectInput("id=division_code", 2, 1, 0, timeWaitAjax),
				new ObjectInput("className=dataUserUse", 2, 1, 0, 0),
				new ObjectInput("className=pcCategory", 0, 1, 0, "pcCategory", 0),
				new ObjectInput("className=pcCode", 2, 2, 0, 0),
				new ObjectInput("className=targeProposal", 2, 1, 0, timeWaitJs),
				new ObjectInput("id=start_maturity_date", 2, 3, 0, timeWaitJs),
				new ObjectInput("id=end_maturity_date", 2, 3, 0, timeWaitJs),
				new ObjectInput("id=reason_maturity", 2, 2, 0, 0),
				new ObjectInput("id=comment", 2, 2, 0, 0),
				new ObjectInput("id=work_place", 2, 2, 0, 0),
				new ObjectInput("id=start_maturity_date", 2, 3, 0, timeWaitJs),
				new ObjectInput("id=end_maturity_date", 2, 3, 0, timeWaitJs),
				new ObjectInput("name=data[User][appointed_place]", 2, 2, 0, 0),
				new ObjectInput("id=appointed_date", 2, 3, 0, timeWaitJs),
				new ObjectInput("className=disable_checkbox", 2, 8, 0, "css=#status_permission_admin", 0),
				new ObjectInput("className=disable_checkbox", 2, 8, 1, "css=#status_access_network", 0),
				new ObjectInput("className=disable_checkbox", 2, 8, 2, "css=#status_limit_usb", 0),
				new ObjectInput("className=disable_checkbox", 2, 8, 3, "css=#status_outside_laptop", 0),
				new ObjectInput("className=disable_checkbox", 2, 8, 4, "css=#status_limit_sd", 0),
				new ObjectInput("className=disable_checkbox", 2, 8, 5, "css=#status_other", 0),
				new ObjectInput("id=note_other", 2, 2, 0, 0),
				new ObjectInput("", 0, 0, 0, "message", 0),
				new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
				new ObjectInput("", 0, 0, 0, "proposals", 0),
		};
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
			login(excel, "Create Proposal", column, array);
		}	
		driver.open(linkStart+linkCreate);
		logger.info("Start fill data to form create");

		driver.autoFill(array, column, excel);
		driver.selectByVisibleText("className=pcCategory", excel.getStringData(column, Common.getIndexFromArray(array, "pcCategory")), 0);
		driver.click("id=buttonSubmit", 200);
		logger.info("End fill data to form create");
		
		//check invalid message
		List<WebElement> ls = driver.getListElenment("className=invalid-msg");
		for(WebElement e : ls){
			if(!e.getText().contentEquals("")){
				logger.error("Ringi messsage error: "+e.getText());
	    		return false;
	    	}
		}	
		if(!driver.getText("className=invalid-msg-note").equals("")){
			logger.error("Ringi message error: "+driver.getText("className=invalid-msg-note"));
    		return false;
		}
		driver.click("id=insertButton", 0);
		
		//write newest id proposal into excel
		ResultSet rs = null;
		rs = DatabaseAction.stmt.executeQuery("select * from pc_proposal order by id DESC limit 1");
		rs.next();
		String proposalId = rs.getString("proposal_id");
		excel.printStringIntoExcel(column, Common.getIndexFromArray(array, "newID"), proposalId);
		excel.write();
		logger.info("start check GUI");
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "message")).equals(driver.getText("id=flashMessage"))
				|| !(linkStart+excel.getStringData(column, Common.getIndexFromArray(array, "urlRedirect"))).equals(driver.getCurrentUrl())){
			logger.error("Error GUI");
			return false;
		}
		logger.info("End check GUI");
		
		if(excel.getStringData(column, Common.getIndexFromArray(array, "chooseValidate")).equals("N")){
			return true;
		}else{
			//Check database
			logger.info("Start check DB");
			int tableProposal = Common.getIndexFromArray(array, "proposals");
			int tablePcProposal = tableProposal + 4;
			int tablePcProposalHistories = tablePcProposal + 33;
			//table proposal
			rs = DatabaseAction.stmt.executeQuery("select * from proposals order by id DESC limit 1");
			rs.next();
			if(db.checkDataInTableProposals(excel, column, tableProposal, rs) == false){
				logger.error("Failed in table proposals");
				return false;
			}
			//table pc_proposal
			rs = DatabaseAction.stmt.executeQuery("select * from pc_proposal order by id DESC limit 1");
			rs.next();
			if(db.checkDataInTablePcProposal(excel, column, tablePcProposal, rs)==false){
				logger.error("Failed in table pc_proposal");
				return false;
			}
			//table pc_proposal_histories
			rs = DatabaseAction.stmt.executeQuery("select * from pc_proposal_histories order by id DESC limit 1");
			rs.next();
			if(db.checkDataInTablePcProposalHistories(excel, column, tablePcProposalHistories, rs) == false){
				logger.error("Failed in table pc_proposal_histories");
				return false;
			}
			logger.info("End check DB");
			return true;
		}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto đề xuất lại đơn PC
	 * @Author: hieuht
	 * @Date: 05/12/2016
	 * @Parametter: 
	 * 		int column: cột chứa data
	 */
	public boolean createProposalAgain(int column) throws Exception {		
		logger=Logger.getLogger("createProposalAgain");

		//Thực hiện reset data nếu dữ liệu trong excel = 2 (đối với mt local)
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
			logger.info("Reset db success");
		}
		
		//Tạo file output excel, truy cập sheet   
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 30);
		excel.accessSheet("Create Proposal Again");
		
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//1: setvisible, 2: sendkey, 3 pickdate, 4 click, 5 chon setvisible theo chi so, 6 click theo chi so
		//ID, action, type, index, name
		ObjectInput[] array = {	
				new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
				new ObjectInput("", 0, 0, 0, "numberUser", 0),
				new ObjectInput("", 0, 0, 0, "idProposal", 0),
				new ObjectInput("id=projectdate", 2, 3, 0, timeWaitJs),
				new ObjectInput("id=checkbox_create_by", 2, 7, 0, "helpBy", timeWaitAjax*2),
				new ObjectInput("id=proposal_change", 2, 1, 0, "userHelpBy", timeWaitAjax),
				new ObjectInput("id=company_code", 2, 1, 0, timeWaitAjax),
				new ObjectInput("id=group_code", 2, 1, 0, timeWaitAjax),
				new ObjectInput("id=division_code", 2, 1, 0, timeWaitAjax),
				new ObjectInput("className=dataUserUse", 2, 1, 0, 0),
				new ObjectInput("className=pcCategory", 0, 1, 0, "pcCategory", 0),
				new ObjectInput("className=pcCode", 2, 2, 0, 0),
				new ObjectInput("className=targeProposal", 2, 1, 0, timeWaitJs),
				new ObjectInput("id=start_maturity_date", 2, 3, 0, timeWaitJs),
				new ObjectInput("id=end_maturity_date", 2, 3, 0, timeWaitJs),
				new ObjectInput("id=reason_maturity", 2, 2, 0, 0),
				new ObjectInput("id=comment", 2, 2, 0, 0),
				new ObjectInput("id=work_place", 2, 2, 0, 0),
				new ObjectInput("id=start_maturity_date", 2, 3, 0, timeWaitJs),
				new ObjectInput("id=end_maturity_date", 2, 3, 0, timeWaitJs),
				new ObjectInput("name=data[User][appointed_place]", 2, 2, 0, 0),
				new ObjectInput("id=appointed_date", 2, 3, 0, timeWaitJs),
				new ObjectInput("className=disable_checkbox", 2, 8, 0, "css=#status_permission_admin", 0),
				new ObjectInput("className=disable_checkbox", 2, 8, 1, "css=#status_access_network", 0),
				new ObjectInput("className=disable_checkbox", 2, 8, 2, "css=#status_limit_usb", 0),
				new ObjectInput("className=disable_checkbox", 2, 8, 3, "css=#status_outside_laptop", 0),
				new ObjectInput("className=disable_checkbox", 2, 8, 4, "css=#status_limit_sd", 0),
				new ObjectInput("className=disable_checkbox", 2, 8, 5, "css=#status_other", 0),
				new ObjectInput("id=note_other", 2, 2, 0, 0),
				new ObjectInput("", 0, 0, 0, "message", 0),
				new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
				new ObjectInput("", 0, 0, 0, "pc_proposal", 0),
		};
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
			login(excel, "Create Proposal Again", column, array);
		}	
		String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
		driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
		logger.info("Start fill data to form create again");
		driver.click("className=proposal_again", timeWaitAjax);

		driver.autoFill(array, column, excel);
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "pcCategory")).equals(""))
			driver.selectByVisibleText("className=pcCategory", excel.getStringData(column, Common.getIndexFromArray(array, "pcCategory")), 0);
		driver.click("id=buttonSubmit", timeWaitJs);
		logger.info("End fill data to form create again");
		//check invalid message
		List<WebElement> ls = driver.getListElenment("className=invalid-msg");
		for(WebElement e : ls){
			if(!e.getText().contentEquals("")){
				logger.error("Ringi message error: "+e.getText());
	    		return false;
	    	}
		}	
		if(!driver.getText("className=invalid-msg-note").equals("")){
			logger.error("Ringi message error: "+driver.getText("className=invalid-msg-note"));
    		return false;
		}
		driver.click("id=insertButton", 0);
		
		logger.info("start check GUI");
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "message")).equals(driver.getText("id=flashMessage"))
				|| !(linkStart+excel.getStringData(column, Common.getIndexFromArray(array, "urlRedirect"))).equals(driver.getCurrentUrl())){
			logger.error("Error GUI");
			return false;
		}
		logger.info("End check GUI");
		
		if(excel.getStringData(column, Common.getIndexFromArray(array, "chooseValidate")).equals("N")){
			return true;
		}else{
			ResultSet rs = null;
			//Check database
			logger.info("Start check DB");
			int tablePcProposal = Common.getIndexFromArray(array, "pc_proposal");
			int tablePcProposalHistories = tablePcProposal + 32;
			int tablePcProposalProcess = tablePcProposalHistories + 6;
			//table pc_proposal
			rs = DatabaseAction.stmt.executeQuery("select * from pc_proposal where proposal_id = "+idProposal);
			rs.next();
			if(db.checkDataInTablePcProposal(excel, column, tablePcProposal, rs) == false){
				logger.error("Failed in table pc_proposal");
				return false;
			}
			
			//table pc_proposal_histories
			rs = DatabaseAction.stmt.executeQuery("select * from pc_proposal_histories order by id DESC limit 1");
			rs.next();
			if(db.checkDataInTablePcProposalHistories(excel, column, tablePcProposalHistories, rs) == false){
				logger.error("Failed in table pc_proposal_histories");
				return false;
			}
			
			//table pc_proposal_process
			rs = DatabaseAction.stmt.executeQuery("select * from pc_proposal_process where proposal_id = " + idProposal);
			while(rs.next()){
				if(!rs.getString("approval_status").equals(excel.getStringData(column, tablePcProposalProcess))){
					logger.error("Failed in table pc_proposal_process");
					return false;
				}
				tablePcProposalProcess+=1;
			}
			logger.info("End check DB");
			return true;
		}
	}

	@Test
	/*
	 * Hàm thực hiện auto setting duyệt đơn PC
	 * @Author: hieuht
	 * @Date: 07/12/2016
	 * @Parametter: 
	 * 		int column: cột chứa data
	 */
	public boolean settingApprove(int column) throws Exception{
		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		//Logger
		logger=Logger.getLogger("settingApprove");
 
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 29);
		excel.accessSheet("Setting Approve");
		
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
								new ObjectInput("", 0, 0, 0, "numberUser", 0),
								new ObjectInput("", 0, 0, 0, "idProposal", 0),
								new ObjectInput("name=lsCategories", 2, 5, 0, timeWaitAjax),
								new ObjectInput("", 2, 4, 0, "use1_group1", 0),
								new ObjectInput("", 2, 4, 0, "use2_group1", 0),
								new ObjectInput("name=lsCategories", 2, 5, 1, timeWaitAjax),
								new ObjectInput("", 2, 4, 0, "use1_group2", 0),
								new ObjectInput("", 2, 4, 0, "use2_group2", 0),
								new ObjectInput("", 0, 1, 0, 0),
								new ObjectInput("", 2, 4, 0, "use_group3", 0),
								new ObjectInput("id=comment", 2, 2, 0, 0),
								new ObjectInput("", 0, 0, 0, "Message", 0),
								new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
								new ObjectInput("", 0, 0, 0, "status_pc_proposal", 0),
								new ObjectInput("", 0, 0, 0, "id_pc_proposal_histories", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "creator_pc_proposal_process", 0)
			};
			if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals(""))
				login(excel, "Setting Approve", column, array);			
			
			String[] arrayQuery = {
					excel.getStringData(column, Common.getIndexFromArray(array, "use1_group1")), 
					excel.getStringData(column, Common.getIndexFromArray(array, "use2_group1")),
					excel.getStringData(column, Common.getIndexFromArray(array, "use1_group2")),
					excel.getStringData(column, Common.getIndexFromArray(array, "use2_group2")),
					excel.getStringData(column, Common.getIndexFromArray(array, "use_group3"))
			};
			
			String[] arrayName = {
					"use1_group1", 
					"use2_group1",
					"use1_group2",
					"use2_group2",
					"use_group3"
			};
			
			String[] arrayIdWeb = {
					"id=Group1User", 
					"id=Group1User",
					"id=Group2User",
					"id=Group2User",
					"id=Group11User"
			};
			
			for(int i=0; i<arrayQuery.length; i++)
				if(!arrayQuery[i].equals(""))
					Common.setIdByUserName(array, arrayName[i], arrayQuery[i], arrayIdWeb[i]);
			
			String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
			driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
			logger.info("Start fill data to form setting approve");
			List<WebElement> list = driver.getListElenment("className=listUser11");
			for(WebElement e : list){
				e.click();
			}
			driver.autoFill(array, column, excel);
			logger.info("End fill data to form setting approve");
			
			//Bấm submit
			try{
				driver.click("id=buttonSubmit", timeWaitJs);
				//Bấm "OK" kết thúc setting duyệt
				driver.click("id=buttonSubmitSetting", 0);
			}catch(Exception e){
				logger.error("Chua nhap form dung yeu cau");
				return false;
			}
			int indexMessage = Common.getIndexFromArray(array, "Message");
			int indexLinkRedirect = Common.getIndexFromArray(array, "urlRedirect");
			//Check GUI
			logger.info("start check GUI");
			if(!excel.getStringData(column, indexMessage).equals(driver.getText("id=flashMessage"))
					|| !(linkStart+excel.getStringData(column, indexLinkRedirect)).equals(driver.getCurrentUrl())){
				logger.error("Error GUI");
				return false;
			}
			logger.info("End check GUI");
			if(excel.getStringData(column, Common.getIndexFromArray(array, "chooseValidate")).equals("N")){
				excel.finish();
				return true;
			}
			else{
				ResultSet rs = null;
				int indexPcProposalProcess = Common.getIndexFromArray(array, "creator_pc_proposal_process");
				int sumColumnProcess = 9;
				int indexTsProposal = Common.getIndexFromArray(array, "status_pc_proposal");
				int indexTsProposalHistories = Common.getIndexFromArray(array, "id_pc_proposal_histories");
				//Check DB			
				logger.info("Start check db");
				//Kiểm tra bảng pc_proposals
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal WHERE proposal_id = "+idProposal);
				rs.next();

				if(!(excel.getStringData(column, indexTsProposal).equals(DatabaseAction.getStringData(rs, "status")))){
					logger.error("Failed in table pc_proposal");
					return false;
				}
	
				//Kiểm tra bảng pc_proposal_histories
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal_histories order by id DESC LIMIT 1");
				rs.next();
				if(!db.checkDataInTablePcProposalHistories(excel, column, indexTsProposalHistories, rs)){
					logger.error("Failed in table pc_proposal_histories");
					return false;
				}
				
				//Kiểm tra bảng pc_proposal_process
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal_process WHERE proposal_id = "+idProposal+" ORDER BY id ASC");
				while(rs.next()){
					if(!db.checkDataInTablePcProposalProcess(excel, column, indexPcProposalProcess, rs)){
						logger.error("Failed in table pc_proposal_process"); 
						return false;
					}
					indexPcProposalProcess += sumColumnProcess;
				}
				logger.info("End check db");
				excel.finish();
				return true;
		}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto setting back đơn PC
	 * @Author: hieuht
	 * @Date: 07/12/2016
	 * @Parametter: 
	 * 		int column: cột chứa data
	 */
	public boolean settingBack(int column) throws Exception{
		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		//Logger
		logger=Logger.getLogger("SettingBackProposal");

		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);   
		excel.accessSheet("Config"); 
		String link = excel.getStringData(1, 29);
		excel.accessSheet("Setting Approve_Back");
		
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
								new ObjectInput("", 0, 0, 0, "numberUser", 0),
								new ObjectInput("", 0, 0, 0, "idProposal", 0),
								new ObjectInput("id=comment", 1, 2, 0, 0),
								new ObjectInput("", 0, 0, 0, "Message", 0),
								new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
								new ObjectInput("", 0, 0, 0, "status_pc_proposal", 0),
								new ObjectInput("", 0, 0, 0, "userIdBack_pc_proposal", 0),
								new ObjectInput("", 0, 0, 0, "pc_proposal_histories", 0),
		};
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals(""))
			login(excel, "Setting Approve_Back", column, array);
			
		String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
		driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
		logger.info("Start fill data to form setting approve back");
		driver.autoFill(array, column, excel);
		logger.info("End fill data to form setting approve");
		
		//Bấm submit
		driver.click("id=buttonBack", timeWaitJs);
		try{
			driver.closeAlertAndGetItsText(timeWaitAjax * 2);
		}catch(Exception e){
			logger.error("Chua nhap form dung yeu cau");
			return false;
		}
		int indexMessage = Common.getIndexFromArray(array, "Message");
		int indexLinkRedirect = Common.getIndexFromArray(array, "urlRedirect");
		//Check GUI
		logger.info("start check GUI");
		if(!excel.getStringData(column, indexMessage).equals(driver.getText("id=flashMessage"))
				|| !(linkStart+excel.getStringData(column, indexLinkRedirect)).equals(driver.getCurrentUrl())){
			logger.error("Error GUI");
			return false;
		}
		logger.info("End check GUI");
		
		if(excel.getStringData(column, Common.getIndexFromArray(array, "chooseValidate")).equals("N")){
			return true;
		}
		else{	
			int indexStatus = Common.getIndexFromArray(array, "status_pc_proposal");
			int indexUserIdBack = Common.getIndexFromArray(array, "userIdBack_pc_proposal");		
			int indexPcProposalHistories = Common.getIndexFromArray(array, "pc_proposal_histories");		
			//Check DB
			ResultSet rs;
			logger.info("Start check db");
			//Kiểm tra bảng pc_proposal
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal WHERE proposal_id = "+idProposal);
			rs.next();
			if(!excel.getStringData(column, indexStatus).equals(DatabaseAction.getStringData(rs, "status"))
					|| !excel.getStringData(column, indexUserIdBack).equals(DatabaseAction.getStringData(rs, "user_id_back"))){
				logger.error("Failed in table pc_proposal");
				return false;
			}
			
			//Kiểm tra bảng pc_proposal_histories
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal_histories order by id DESC LIMIT 1");
			rs.next();
			if(!db.checkDataInTablePcProposalHistories(excel, column, indexPcProposalHistories, rs)){
				logger.error("Failed in table pc_proposal_histories");
				return false;
			}
			logger.info("End check db");
			return true;
		}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto duyệt đơn PC
	 * @Author: hieuht
	 * @Date: 09/12/2016
	 * @Parametter: 
	 * 		int column: cột chứa data
	 */
	public boolean approve(int column) throws Exception {
		//Logger
		logger=Logger.getLogger("ApproveProposal");
		
		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 31);
		excel.accessSheet("Approve");

		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, index, name
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
				new ObjectInput("", 0, 0, 0, "numberUser", 0),
				new ObjectInput("", 0, 0, 0, "idProposal", 0),
				new ObjectInput("", 0, 0, 0, "choose", 0),
				new ObjectInput("id=content_proposal", 1, 2, 0, "comment", 0),
				new ObjectInput("", 0, 0, 0, "userName1", 0),
				new ObjectInput("", 0, 0, 0, "password1", 0),
				new ObjectInput("", 0, 0, 0, "choose1", 0),
				new ObjectInput("id=content_proposal", 0, 2, 0, "comment1", 0),
				new ObjectInput("", 0, 0, 0, "userName2", 0),
				new ObjectInput("", 0, 0, 0, "password2", 0),
				new ObjectInput("id=content_proposal", 0, 2, 0, "comment2", 0),
				new ObjectInput("", 0, 0, 0, "Message", 0),
				new ObjectInput("", 0, 0, 0, "urlRedirect1", 0),
				new ObjectInput("", 0, 0, 0, "urlRedirect2", 0),
				new ObjectInput("", 0, 0, 0, "urlRedirect3", 0),
				new ObjectInput("", 0, 0, 0, "status_pc", 0),
				new ObjectInput("", 0, 0, 0, "pc_proposal_histories", 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, "approval_status", 0),
				new ObjectInput("", 0, 0, 0, "status_all", 0),
				new ObjectInput("", 0, 0, 0, "comment_ts", 0),
				};
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals(""))
			login(excel, "Approve", column, array);
		
		String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
		excel.printStringIntoExcel(column, Common.getIndexFromArray(array, "idProposal"), idProposal);
		String alert = "";			
		logger.info("Start fill data to form approve proposal");
		driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
		
		String linkRedirect1 = "";
		String linkRedirect2 = "";
		String linkRedirect3 = "";
		String messageRedirect1 = "";
		String messageRedirect2 = "";
		String messageRedirect3 = "";
		int indexChoose = Common.getIndexFromArray(array, "choose");
		if(excel.getStringData(column, indexChoose).equals("1") || excel.getStringData(column, indexChoose).equals("2")){
			if(excel.getStringData(column, indexChoose).equals("1")){
				driver.getElementFormList("id=chkSelected", 0).click();
			}
			else if(excel.getStringData(column, indexChoose).equals("2")){
				driver.getElementFormList("id=chkSelected", 1).click();
			}
			driver.sendKey("id=content_proposal", excel.getStringData(column, Common.getIndexFromArray(array, "comment")), 0);
			driver.click("id=approved_proposal", timeWaitJs);
			alert = driver.closeAlertAndGetItsText(timeWaitAjax * 2);
			if(!alert.equals("承認します。よろしいですか？")){
				logger.info("Error in arlert");
				return false;
			}	
			linkRedirect1 = driver.getCurrentUrl();
			messageRedirect1 = driver.getText("id=flashMessage");
			
			if(!excel.getStringData(column, Common.getIndexFromArray(array, "userName1")).equals("")){
			    driver.login(linkStart+linkLogin, "id=UserUsername", excel.getStringData(column, Common.getIndexFromArray(array, "userName1")), "id=UserPassword", excel.getStringData(column, Common.getIndexFromArray(array, "password1")), "css=input[type=\"submit\"]");
			    driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
			    
			    int indexChoose1 = Common.getIndexFromArray(array, "choose1");

			    if(excel.getStringData(column, indexChoose1).equals("1")){
					driver.getElementFormList("id=chkSelected", 0).click();
				}
				else if(excel.getStringData(column, indexChoose1).equals("2")){
					driver.getElementFormList("id=chkSelected", 1).click();
				}
				driver.sendKey("id=content_proposal", excel.getStringData(column, Common.getIndexFromArray(array, "comment1")), 0);
				driver.click("id=approved_proposal", timeWaitJs);
				alert = driver.closeAlertAndGetItsText(timeWaitAjax * 2);
				if(!alert.equals("承認します。よろしいですか？")){
					logger.info("Error in arlert");
					return false;
				}	
				messageRedirect2 = driver.getText("id=flashMessage");
				linkRedirect2 = driver.getCurrentUrl();
			}else{
				linkRedirect2 = linkStart;
				messageRedirect2 = messageRedirect1;
			}
		}else if(excel.getStringData(column, indexChoose).equals("3")){
			driver.getElementFormList("id=chkSelected", 0).click();
			driver.getElementFormList("id=chkSelected", 1).click();
			driver.sendKey("id=content_proposal", excel.getStringData(column, Common.getIndexFromArray(array, "comment")), 0);
			driver.click("id=approved_proposal", 0);
			alert = driver.closeAlertAndGetItsText(timeWaitAjax * 2);
			if(!alert.equals("承認します。よろしいですか？")){
				logger.info("Error in arlert");
				return false;
			}	
			messageRedirect1 = driver.getText("id=flashMessage");
			messageRedirect2 = messageRedirect1;
			linkRedirect1 = driver.getCurrentUrl();
			linkRedirect2 = linkStart;
		}	
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "userName2")).equals("")){
		    driver.login(linkStart+linkLogin, "id=UserUsername", excel.getStringData(column, Common.getIndexFromArray(array, "userName2")), "id=UserPassword", excel.getStringData(column, Common.getIndexFromArray(array, "password2")), "css=input[type=\"submit\"]");
		    
		    driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
		    driver.click("id=chkSelected", 0);
		    driver.sendKey("id=content_proposal", excel.getStringData(column, Common.getIndexFromArray(array, "comment2")), 0);
		    driver.click("id=approved_proposal", timeWaitJs);
		    alert = driver.closeAlertAndGetItsText(timeWaitAjax * 2);
			if(!alert.equals("承認します。よろしいですか？")){
				logger.info("Error in arlert");
				return false;
			}	
			linkRedirect3 = driver.getCurrentUrl();
			messageRedirect3 = driver.getText("id=flashMessage");
		}
		else{
			linkRedirect3 = linkStart;
			messageRedirect3 = messageRedirect1;
		}
		logger.info("End fill data to form approve proposal");
		
		int indexMessage = Common.getIndexFromArray(array, "Message");
		int indexLinkRedirect1 = Common.getIndexFromArray(array, "urlRedirect1");
		int indexLinkRedirect2 = Common.getIndexFromArray(array, "urlRedirect2");
		int indexLinkRedirect3 = Common.getIndexFromArray(array, "urlRedirect3");
		//Check GUI
		logger.info("start check GUI");
		String linkSpecial = excel.getStringData(column, indexLinkRedirect1);
		if(linkSpecial.endsWith("id:0")){
			linkSpecial = linkSpecial.substring(0, linkSpecial.length()-1).concat(idProposal);
		}
		if(!excel.getStringData(column, indexMessage).equals(messageRedirect1)
				|| !excel.getStringData(column, indexMessage).equals(messageRedirect2)
				|| !excel.getStringData(column, indexMessage).equals(messageRedirect3)
				|| !(linkStart+linkSpecial).equals(linkRedirect1)
				|| !(linkStart+excel.getStringData(column, indexLinkRedirect2)).equals(linkRedirect2)
				|| !(linkStart+excel.getStringData(column, indexLinkRedirect3)).equals(linkRedirect3)){
			logger.error("Error GUI");
			return false;
		}
		logger.info("End check GUI");
		
		if(excel.getStringData(column, Common.getIndexFromArray(array, "chooseValidate")).equals("N")){
			excel.finish();
			return true;
		}else{
			int indexPcProposalProcess1 = Common.getIndexFromArray(array, "approval_status");
			int indexPcProposalProcess2 = Common.getIndexFromArray(array, "status_all");
			int indexPcProposalProcess3 = Common.getIndexFromArray(array, "comment_ts");
			int sumProcess = 3;
			int indexStatus = Common.getIndexFromArray(array, "status_pc");
			int indexTsProposalHistories = Common.getIndexFromArray(array, "pc_proposal_histories");
			
			//Check db
			ResultSet rs;
			//Kiểm tra bảng pc_proposal_histories
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal_histories order by id DESC LIMIT 1");
			rs.next();
			if(!db.checkDataInTablePcProposalHistories(excel, column, indexTsProposalHistories, rs)){
				logger.error("Failed in table pc_proposal_histories");
				return false;
			}
			
			//Kiểm tra bảng pc_proposals
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal WHERE proposal_id = "+idProposal);
			rs.next();
			if(!(excel.getStringData(column, indexStatus).equals(DatabaseAction.getStringData(rs, "status")))){
				logger.error("Failed in table pc_proposal");
				return false;
			}
			
			//Kiểm tra bảng pc_proposal_process
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal_process WHERE proposal_id = "+idProposal+" ORDER BY id ASC");
			while(rs.next()){
				if(!(excel.getStringData(column, indexPcProposalProcess1).equals(DatabaseAction.getStringData(rs, "approval_status")))
					|| !(excel.getStringData(column, indexPcProposalProcess2).equals(DatabaseAction.getStringData(rs, "status_all")))
					|| !(excel.getStringData(column, indexPcProposalProcess3).equals(DatabaseAction.getStringData(rs, "comment")))){
					logger.error("Failed in table pc_proposal_process");
					return false;
				}
				indexPcProposalProcess1+= sumProcess;
				indexPcProposalProcess2+= sumProcess;
				indexPcProposalProcess3+= sumProcess;
			}
			logger.info("End check db");
			excel.finish();
			return true;
		}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto back đơn PC
	 * @Author: hieuht
	 * @Date: 14/12/2016
	 * @Parametter:
	 * 		int column: cột chứa dữ liệu
	 */
	public boolean back(int column) throws Exception{
		//Logger
		logger=Logger.getLogger("Back Proposal");
	
		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);  
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 31);
		excel.accessSheet("Back");
		
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, index, name
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
								new ObjectInput("", 0, 0, 0, "numberUser", 0),
								new ObjectInput("", 0, 0, 0, "idProposal", 0),
								new ObjectInput("", 0, 0, 0, "choose", 0),
								new ObjectInput("id=content_proposal", 1, 2, 0, "comment", 0),
								new ObjectInput("", 0, 0, 0, "message", 0),
								new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
								new ObjectInput("", 0, 0, 0, "idPcProposalHistories", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "approval_status", 0),
								new ObjectInput("", 0, 0, 0, "status_all", 0),
								new ObjectInput("", 0, 0, 0, "comment_ts", 0),
				};
			if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals(""))
				login(excel, "Back", column, array);	
		
			String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
			String alert = "";	
			logger.info("Start fill data to form back proposal");
			driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
			
			int indexChoose = Common.getIndexFromArray(array, "choose");
			if(excel.getStringData(column, indexChoose).equals("1")){
				array[indexChoose].setActon(1);
				array[indexChoose].setId("id=chkSelected");
				array[indexChoose].setIndex(0);
				array[indexChoose].setType(6);
				driver.autoFill(array, column, excel);
			}else if(excel.getStringData(column, indexChoose).equals("2")){
				array[indexChoose].setActon(1);
				array[indexChoose].setId("id=chkSelected");
				array[indexChoose].setIndex(1);
				array[indexChoose].setType(6);
				driver.autoFill(array, column, excel);
			}else if(excel.getStringData(column, indexChoose).equals("3")){
				driver.getElementFormList("id=chkSelected", 0).click();
				driver.getElementFormList("id=chkSelected", 1).click();
				driver.sendKey("id=content_proposal", excel.getStringData(column, Common.getIndexFromArray(array, "comment")), 0);
			}
			//Kiểm tra thành công
			driver.click("id=back_proposal", timeWaitJs);
			logger.info("End fill form");
			try{
				alert = driver.closeAlertAndGetItsText(timeWaitAjax * 2);
			}catch(Exception e){
				logger.error("Chua nhap form dung yeu cau");
				return false;
			}
			if(!alert.equals("差し戻します。よろしいですか？")){
				logger.error("Ringi alert error: "+alert);
				return false;
			}	
			if(!driver.getText("id=flashMessage").equals("差戻に成功しました。")){
				logger.error("Back don that bai");
				return false;
			}	
			logger.info("Back don thanh cong");
			
			int indexLinkRedirect = Common.getIndexFromArray(array, "urlRedirect");
			int indexMessage = Common.getIndexFromArray(array, "message");
			//Check GUI
			logger.info("start check GUI");
			if(!excel.getStringData(column, indexMessage).equals(driver.getText("id=flashMessage"))
					|| !(linkStart+excel.getStringData(column, indexLinkRedirect)).equals(driver.getCurrentUrl())){
				logger.error("Error GUI");
				return false;
			}
			logger.info("End check GUI");
			
			if(excel.getStringData(column, Common.getIndexFromArray(array, "chooseValidate")).equals("N")){
				return true;
			}else{
				int indexPcProposalProcess1 = Common.getIndexFromArray(array, "approval_status");
				int indexPcProposalProcess2 = Common.getIndexFromArray(array, "status_all");
				int indexPcProposalProcess3 = Common.getIndexFromArray(array, "comment_ts");
				int sumProcess = 3;		
				int indexPcProposalHistories = Common.getIndexFromArray(array, "idPcProposalHistories");
				
				//Check db
				ResultSet rs;
				//Kiểm tra bảng pc_proposal_histories
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal_histories order by id DESC LIMIT 1");
				rs.next();
				if(!db.checkDataInTablePcProposalHistories(excel, column, indexPcProposalHistories, rs)){
					logger.error("Failed in table pc_proposal_histories");
					return false;
				}
				
				//Kiểm tra bảng pc_proposal_process
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal_process WHERE proposal_id = "+idProposal+" ORDER BY id ASC");
				while(rs.next()){
					if(!(excel.getStringData(column, indexPcProposalProcess1).equals(DatabaseAction.getStringData(rs, "approval_status")))
						|| !(excel.getStringData(column, indexPcProposalProcess2).equals(DatabaseAction.getStringData(rs, "status_all")))
						|| !(excel.getStringData(column, indexPcProposalProcess3).equals(DatabaseAction.getStringData(rs, "comment")))){
						logger.error("Failed in table pc_proposal_process");
						return false;
					}
					indexPcProposalProcess1+= sumProcess;
					indexPcProposalProcess2+= sumProcess;
					indexPcProposalProcess3+= sumProcess;
				}
				logger.info("End check db");
				return true;
			}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto cancel đơn PC
	 * @Author: hieuht
	 * @Date: 14/12/2016
	 * @Parametter:
	 * 		int column: cột chứa dữ liệu
	 */
	public boolean cancel(int column) throws Exception{
		//Logger
		logger=Logger.getLogger("Cancel Proposal");

		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);  
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 31);
		excel.accessSheet("Cancel Approve");
		
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, index, name
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
				new ObjectInput("", 0, 0, 0, "numberUser", 0),
				new ObjectInput("", 0, 0, 0, "idProposal", 0),
				new ObjectInput("", 0, 0, 0, "choose", 0),
				new ObjectInput("id=content_proposal", 1, 2, 0, "comment", 0),
				new ObjectInput("", 0, 0, 0, "message", 0),
				new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
				new ObjectInput("", 0, 0, 0, "idPcProposalHistories", 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, "approval_status", 0),
				new ObjectInput("", 0, 0, 0, "comment_ts", 0),
				};
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals(""))
			login(excel, "Cancel Approve", column, array);
		
		String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
		String alert = "";	
		logger.info("Start fill data to form cancel proposal");
		driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
		
		int indexChoose = Common.getIndexFromArray(array, "choose");
		if(excel.getStringData(column, indexChoose).equals("1")){
			array[indexChoose].setActon(1);
			array[indexChoose].setId("id=chkSelected");
			array[indexChoose].setIndex(0);
			array[indexChoose].setType(6);
			driver.autoFill(array, column, excel);
		}else if(excel.getStringData(column, indexChoose).equals("2")){
			array[indexChoose].setActon(1);
			array[indexChoose].setId("id=chkSelected");
			array[indexChoose].setIndex(1);
			array[indexChoose].setType(6);
			driver.autoFill(array, column, excel);
		}else if(excel.getStringData(column, indexChoose).equals("3")){
			driver.getElementFormList("id=chkSelected", 0).click();
			driver.getElementFormList("id=chkSelected", 1).click();
			driver.sendKey("id=content_proposal", excel.getStringData(column, Common.getIndexFromArray(array, "comment")), 0);
		}
		//Kiểm tra thành công
		driver.click("id=cancel_proposal", timeWaitJs);
		logger.info("End fill form");
		alert = driver.closeAlertAndGetItsText(timeWaitAjax * 2);
		if(!alert.equals("承認取消しします。よろしいですか？")){
			logger.error("Ringi alert error: "+alert);
			return false;
		}	
		logger.info("Cancel don thanh cong");
		
		int indexMessage = Common.getIndexFromArray(array, "message");
		int indexLinkRedirect = Common.getIndexFromArray(array, "urlRedirect");
		//Check GUI
		logger.info("start check GUI");
		String Link = excel.getStringData(column, indexLinkRedirect);
		if(Link.endsWith("id:0"))
			Link = Link.substring(0, Link.length()-1).concat(idProposal);
		if(!excel.getStringData(column, indexMessage).equals(driver.getText("id=flashMessage"))
				|| !(linkStart+Link).equals(driver.getCurrentUrl())){
			logger.error("Error GUI");
			return false;
		}
		logger.info("End check GUI");
		
		if(excel.getStringData(column, Common.getIndexFromArray(array, "chooseValidate")).equals("N")){
			return true;
		}else{
			int indexPcProposalProcess1 = Common.getIndexFromArray(array, "approval_status");
			int indexPcProposalProcess3 = Common.getIndexFromArray(array, "comment_ts");
			int sumProcess = 2;			
			int indexPcProposalHistories = Common.getIndexFromArray(array, "idPcProposalHistories");
			
			//Check db
			ResultSet rs;
			//Kiểm tra bảng pc_proposal_histories
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal_histories order by id DESC LIMIT 1");
			rs.next();
			if(!db.checkDataInTablePcProposalHistories(excel, column, indexPcProposalHistories, rs)){
				logger.error("Failed in table pc_proposal_histories");
				return false;
			}
			
			//Kiểm tra bảng pc_proposal_process
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal_process WHERE proposal_id = "+idProposal+" ORDER BY id ASC");
			while(rs.next()){
				if(!(excel.getStringData(column, indexPcProposalProcess1).equals(DatabaseAction.getStringData(rs, "approval_status")))
					|| !(excel.getStringData(column, indexPcProposalProcess3).equals(DatabaseAction.getStringData(rs, "comment")))){
					logger.error("Failed in table pc_proposal_process");
					return false;
				}
				indexPcProposalProcess1+= sumProcess;
				indexPcProposalProcess3+= sumProcess;
			}
			logger.info("End check db");
			return true;
		}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto quản lý task đơn PC
	 * @Author: hieuht
	 * @Date: 14/12/2016
	 * @Parametter:
	 * 		int column: cột chứa dữ liệu
	 */
	public boolean statusManagement(int column) throws Exception{
		//Logger
		logger=Logger.getLogger("Determine Proposal");

		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		logger.info("Read data form excel");
		
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild); 
		excel.accessSheet("Config");
		
		//Array link from sheet Config
		ObjectInput[] arrayLink = {
							new ObjectInput(excel.getStringData(0, 33), excel.getStringData(1, 33)),
							new ObjectInput(excel.getStringData(0, 34), excel.getStringData(1, 34)),
							new ObjectInput(excel.getStringData(0, 35), excel.getStringData(1, 35)),
							new ObjectInput(excel.getStringData(0, 36), excel.getStringData(1, 36)),
							new ObjectInput(excel.getStringData(0, 37), excel.getStringData(1, 37)),
							new ObjectInput(excel.getStringData(0, 38), excel.getStringData(1, 38)),
		};
		String link = "";
		excel.accessSheet("Status Management");
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, index, name
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
								new ObjectInput("", 0, 0, 0, "numberUser", 0),
								new ObjectInput("", 0, 0, 0, "idProposal", 0),
								new ObjectInput("", 0, 0, 0, "typeLink", 0),
								new ObjectInput("className=status_pc", 2, 1, 0, 0),
								new ObjectInput("id=execute_date", 2, 3, 0, timeWaitJs),
								new ObjectInput("id=input_comment", 2, 2, 0, "comment", 0),
								new ObjectInput("", 0, 0, 0, "message", 0),
								new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
								new ObjectInput("", 0, 0, 0, "status", 0),
								new ObjectInput("", 0, 0, 0, "last_operator", 0),
								new ObjectInput("", 0, 0, 0, "user_id_change_complete", 0),
								new ObjectInput("", 0, 0, 0, "execute_date", 0),
								new ObjectInput("", 0, 0, 0, "idPcProposalHistories", 0),
								};
			if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals(""))
				login(excel, "Status Management", column, array);

			String typeLink = excel.getStringData(column, Common.getIndexFromArray(array, "typeLink"));
			for(int i=0; i<arrayLink.length; i++){
				if(arrayLink[i].getName().equals(typeLink)){
					link = arrayLink[i].getId();
				}
			}
			String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
			logger.info("Start fill data to form status management proposal");
			driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
			driver.autoFill(array, column, excel);
			driver.click("id=writeProposal", timeWaitJs);
			try{
				driver.closeAlertAndGetItsText((int)(timeWaitAjax * 3));
			} catch (Exception e) {
				logger.error("Chua nhap form dung yeu cau");
			}
			int indexLinkRedirect = Common.getIndexFromArray(array, "urlRedirect");
			int indexMessage = Common.getIndexFromArray(array, "message");
			//Check GUI
			logger.info("start check GUI");
			if(!excel.getStringData(column, indexMessage).equals(driver.getText("id=flashMessage"))
					|| !(linkStart+excel.getStringData(column, indexLinkRedirect)).equals(driver.getCurrentUrl())){
				logger.error("Error GUI");
				return false;
			}
			logger.info("End check GUI");
			
			if(excel.getStringData(column, Common.getIndexFromArray(array, "chooseValidate")).equals("N")){
				return true;
			}else{
				int status = Common.getIndexFromArray(array, "status");
				int last_operator = Common.getIndexFromArray(array, "last_operator");
				int user_id_change_complete = Common.getIndexFromArray(array, "user_id_change_complete");
				int execute_date = Common.getIndexFromArray(array, "execute_date");
				int pc_histories = Common.getIndexFromArray(array, "idPcProposalHistories");
				//Check db
				ResultSet rs;
				rs=null;
				logger.info("Start check db");
				//Kiểm tra bảng pc_proposal_histories
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal_histories order by id DESC LIMIT 1");
				rs.next();
				if(!db.checkDataInTablePcProposalHistories(excel, column, pc_histories, rs)){
					logger.error("Failse in table pc_proposal_histories");
					return false;
				}
				
				//Kiểm tra bảng pc_proposals
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM pc_proposal WHERE proposal_id = "+idProposal);
				rs.next();
				if(!(excel.getStringData(column, status).equals(DatabaseAction.getStringData(rs, "status")))
						|| !(excel.getStringData(column, last_operator).equals(DatabaseAction.getStringData(rs, "last_operator")))
						|| !(excel.getStringData(column, user_id_change_complete).equals(DatabaseAction.getStringData(rs, "user_id_change_complete")))
						|| !(excel.getStringData(column, execute_date).equals(DatabaseAction.getStringData(rs, "execute_date")))){
					logger.error("Failse in table pc_proposal");
					return false;
				}
				logger.info("End check db");
				return true;
			}
	}
	
	/*
	 * author: hieuht
	 * login
	 * Parameter:
	 * 		ExcelAction excel: file excel
	 * 		String nameCurrentSheet: sheet hiện tại
	 * 		int column: chỉ số cột lấy dữ liệu
	 * 		ObjectInput[] array: mảng lấy chỉ số dòng
	 */
	private void login(ExcelAction excel, String nameCurrentSheet, int column, ObjectInput[] array ) throws Exception{
		int numberColumnUserName = 0;
		int numberColumnPassword = 1;
		excel.accessSheet(nameCurrentSheet);
		int rowGetUsername = Integer.parseInt(excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")))-1;
		int rowGetPassword = Integer.parseInt(excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")))-1;
		excel.accessSheet("Login");
		String userName = excel.getStringData(numberColumnUserName, rowGetUsername);
		String password = excel.getStringData(numberColumnPassword, rowGetPassword);
		driver.login(linkStart+linkLogin, "id=UserUsername", userName, "id=UserPassword", password, "css=input[type=\"submit\"]");
		excel.accessSheet(nameCurrentSheet);
	}
	
	@AfterTest
	/*
	 * Hàm thực hiện đóng trình duyệt FireFox và mở file excel
	 * @Author: hieuht
	 * @Date:01/12/2016
	 */
	public void closeBrowser() throws Exception{
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);  
		excel.accessSheet("Create Proposal");
		for(int i=4; i<256; i++)
			excel.printStringIntoExcel(i, 2, "");
		excel.finish();
		logger=Logger.getLogger("closeBrower");
		driver.close();
		logger.info("Close firefox");
		String pathOut = File.substring(0, File.length()-4)+"_Output.xls";
		//Runtime.getRuntime().exec("cmd /c start " + pathOut);
		logger.info("Open file successfully");
		endTime = System.currentTimeMillis();
		long totalTime = endTime - startTime;
		logger.info("Tong thoi gian chay: "+ totalTime/1000 + "s ~ "+totalTime/1000/60+"p"+totalTime/1000%60+"s");
		logger.info("Tong time nghi: "+ driver.getCountTime()/1000 + "s ~ "+driver.getCountTime()/1000/60+"p"+driver.getCountTime()/1000%60+"s");
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		Date date = new Date();
		logger.info("End autotest don PC at "+ dateFormat.format(date));
	}
}


