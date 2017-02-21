/*
 * Lớp thực hiện các chức năng autotest với đơn đi lại
 * @Author: hieuht
 * @Date: 16/09/2016
 */
package DonDiLai;

import java.awt.Desktop;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.Set;

import javax.swing.JOptionPane;

import org.apache.log4j.FileAppender;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.ElementNotVisibleException;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.jcraft.jsch.JSch;
import com.jcraft.jsch.Session;
import com.mysql.jdbc.jdbc2.optional.MysqlDataSource;

import DonDiLai.CheckDatabase;
import Common.*;
import jxl.write.WriteException;

public class FunctionTest extends FileAppender{
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
	final String File = "DataTest\\DonDiLai\\Data-DonDiLai.xls";
	public CheckDatabase db = new CheckDatabase();
	public int resetData = -1;
	private String dbNameDelete = "";
	private String dbNameImport = "";
	private String linkStart = "";
	private String linkLogin;
	private String environment = "";
	public int timeWaitAjax = 0;
	private int timeWaitJs = 200;
	private long startTime;
	private long endTime;
	private String buildId;
	private String timeBuild;
	private String[] flow;
	Logger logger = null;
	/*
	 * Hàm thực hiện mở trình duyệt FireFox version 46 và đăng nhập vào Ringi
	 * @Author: hieuht
	 * @Date:16/09/2016
	 */
	public void openBrowser() throws Exception {
		DateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH.mm");
		Date frmDate = sdf.parse(timeBuild);

		DateFormat sdff = new SimpleDateFormat("yyyy-MM-dd HH.mm");
	    System.setProperty("current.date", sdff.format(frmDate));
	    System.setProperty("current.proposal", "DonDiLai");
	    System.setProperty("current.build", "#"+buildId);
		//Khởi tạo logger
		logger=Logger.getLogger("openBrower");
		//Config log4j
		PropertyConfigurator.configure("log4j.properties");
		
		Common.readExcelDefine();
		
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		Date date = new Date();
		logger.info("Start autotest don di lai at " + dateFormat.format(date));
		startTime = System.currentTimeMillis();
		//Truy cập vào file input excel sheet 0, đọc các dữ liệu config
		ExcelAction excel = new ExcelAction();
		excel.createFileOutput(File, buildId, timeBuild);
		excel.accessSheet("Config");
		int column = 1;		
		logger.info("Read data config");
//		linkStart = excel.getStringData(column, 6);
//		pathFile = excel.getStringData(column, 10);
//		timeWaitAjax = Integer.parseInt(excel.getStringData(column, 9));
		resetData = Integer.parseInt(excel.getStringData(column, 18));
		dbNameImport = excel.getStringData(column, 19);
		dbNameDelete = excel.getStringData(column, 20);
		linkLogin = excel.getStringData(column, 2);
//		userName = excel.getStringData(column, 3);
//		password = excel.getStringData(column, 4);
//		environment = excel.getStringData(1, 7);
		//Tạo file output, sheet 0 
		excel.accessSheet("Config");
		
		DatabaseAction.setInputFile(File);
		DatabaseAction.doSshTunnel(environment, 0, buildId, timeBuild);
		logger.info("Autotest connect db successfully");
		if(resetData == 1 || resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
			logger.info("Reset db successfully");
		}
	    //Mở trình duyệt Firefox
	    driver = new WebDriverAction();
	    logger.info("Open firefox brower");
	    
	    //Mở link, điền thông tin: username, password rồi đăng nhập
//	    logger.info("Fill data to form login");
//	    driver.open(linkStart+linkLogin, timeWaitAjax);
//	    driver.login("id=UserUsername", userName, "id=UserPassword", password, "css=input[type=\"submit\"]");
//	    logger.info("Login successfully");
	    //Đóng file input excel, ghi thay đổi và đóng file output excel
	    try{
	    	excel.finish();
	    }catch(FileNotFoundException e){
	    	driver.close();
	    	JOptionPane.showMessageDialog(null, "Chưa đóng file excel");
	    	logger.info("Please close file excel");
	    	System.exit(0);
	    }
	}
		
	@Test
	/*
	 * Start
	 * @Author: hieuht
	 * @Date: 02/1/2016
	 */
	public void start() throws Exception{
		try {
		openBrowser();
		//Khởi tạo logger
		logger=Logger.getLogger("start");
		//Tạo file output excel, truy cập sheet 
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);  
		excel.accessSheet("Flow");
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "testcase", 0),
				new ObjectInput("", 0, 0, 0, "createProposal", 0),
				new ObjectInput("", 0, 0, 0, "settingApprove", 0),
				new ObjectInput("", 0, 0, 0, "settingBack", 0),
				new ObjectInput("", 0, 0, 0, "approve", 0),
				new ObjectInput("", 0, 0, 0, "back", 0),
				new ObjectInput("", 0, 0, 0, "cancel", 0),
				new ObjectInput("", 0, 0, 0, "reject", 0),
				new ObjectInput("", 0, 0, 0, "determine", 0),
				new ObjectInput("", 0, 0, 0, "edit", 0),
				new ObjectInput("", 0, 0, 0, "createAgain", 0),
				new ObjectInput("", 0, 0, 0, "result", 0),
				new ObjectInput("", 0, 0, 0, "type", 0)};
				
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
		} catch (Exception e) {
			logger.error(e.getMessage());
		}
	}
	
	/*
	 * author: hieuht
	 * Call các hàm chức năng
	 * Parameter: 
	 * 			ExcelAction excel: file excel đọc data
	 * 			ObjectInput[] array: mảng xác định vị trí cột lấy data
	 * 			int row: dòng chứa data
	 */
	public boolean doFlow(ExcelAction excel, ObjectInput[] array, int row)throws Exception{
		for(int i=1; i<array.length; i++){
			if(!excel.getStringData(Common.getIndexFromArray(array, array[i].getName()), row).equals("")){
			int column = excel.getColumn(excel.getStringData(Common.getIndexFromArray(array, array[i].getName()), row));
			DatabaseAction.doSshTunnel(environment, 1, buildId, timeBuild);
			switch(array[i].getName()){
				case "createProposal":{
					if(createProposal(column) == false)
						return false;
					else
						break;
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
				case "reject":{
					if(reject(column) == false)
						return false;
					else
						break;
				}
				case "determine":{
					if(determine(column) == false)
						return false;
					else
						break;
				}
				case "edit":{
					if(edit(column) == false)
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
			}
		}
		return true;
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
	 * Hàm thực hiện auto tạo đơn đi lại
	 * @Author: hieuht
	 * @Date: 16/09/2016
	 */
	public boolean createProposal(int column) throws Exception {
		//Khởi tạo logger
		logger=Logger.getLogger("createMultiProposal");
	 
		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		logger.info("Read data form excel");
		//Tạo file output excel, truy cập sheet 
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);   
		excel.accessSheet("Config");
		String linkCreate = excel.getStringData(1, 28);
		excel.accessSheet("Create Proposal");
		
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, setActionChild, actionChild
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
								new ObjectInput("", 0, 0, 0, "numberUser", 0),
								new ObjectInput("", 0, 0, 0, "newID", 0),
								new ObjectInput("id=UserCreateProposal", 0, 0, 0, "userHelp", 0),
								new ObjectInput("", 0, 0, 0, "idUserHelp", 0),
								new ObjectInput("id=company_code", 2, 1, 0, timeWaitAjax),
								new ObjectInput("id=group_code", 2, 1, 0, timeWaitAjax),
								new ObjectInput("id=division_code", 2, 1, 0, timeWaitAjax),
								new ObjectInput("id=type_id", 2, 1, 0, 0),
								new ObjectInput("id=category_id", 2, 1, 0, 0),
								new ObjectInput("name=data[User][change_date]", 2, 3, 0, timeWaitJs),
								new ObjectInput("", 0, 4, 0, "supplyCost", timeWaitJs),
								new ObjectInput("id=reason_type", 2, 1, 0, 0),
								new ObjectInput("id=post_code", 2, 2, 0, timeWaitAjax),
								new ObjectInput("id=address", 2, 2, 0, 0),
								new ObjectInput("id=route_names_1", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-1", 2, 2, 0, timeWaitAjax),
								new ObjectInput("id=ts-distance-to-1", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_1", 2, 2, 0, 0),
								new ObjectInput("id=route_names_2", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-2", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-2", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_2", 2, 2, 0, 0),
								new ObjectInput("id=route_names_3", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-3", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-3", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_3", 2, 2, 0, 0),
								new ObjectInput("id=route_names_4", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-4", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-4", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_4", 2, 2, 0, 0),
								new ObjectInput("id=route_names_5", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-5", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-5", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_5", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_one_day", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_one_month", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_three_month", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_six_month", 2, 2, 0, 0),
								new ObjectInput("id=comment", 2, 2, 0, 0),
								new ObjectInput("id=origin-input", 2, 2, 0, 0),
								new ObjectInput("id=destination-input", 2, 2, 0, timeWaitJs),
								new ObjectInput("", 0, 0, 0, "Message", 0),
								new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
								new ObjectInput("", 0, 0, 0, "id_proposal", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "id_ts_proposal", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0), 
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "id_ts_proposal_histories", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "id_ts_proposal_route", 0),
		};
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
			login(excel, "Create Proposal", column, array);
		}	
		
		int sumColumnRoute = 5;
		int supplyCost = Common.getIndexFromArray(array, "supplyCost");
		if(excel.getStringData(column, supplyCost).equals("1")){
			array[supplyCost].setId("id=supply_cost_yes");
			array[supplyCost].setActon(1);
		}else if(excel.getStringData(column, supplyCost).equals("0")){
			array[supplyCost].setId("id=supply_cost_no");
			array[supplyCost].setActon(1);
		}
		driver.open(linkStart+linkCreate);
		driver.sleep(timeWaitAjax);
		int helpBy = Common.getIndexFromArray(array, "userHelp");
		int idHelpBy = Common.getIndexFromArray(array, "idUserHelp");
		if(!excel.getStringData(column, helpBy).equals("")){
			driver.sendKey("id=UserCreateProposal", excel.getStringData(column, helpBy), timeWaitAjax*2);
			try{
				ResultSet rs = null;
				rs = DatabaseAction.stmt.executeQuery("SELECT id FROM users where username = "+excel.getStringData(column, idHelpBy));
				rs.next();
				driver.click("xpath=//li[@class='"+rs.getString("id")+"']", timeWaitAjax);
			} catch (Exception e) {
				logger.error("Khong tim thay user tao thay");
				return false;
			}
		}
		
		logger.info("Start fill data to form create");
		int routeCount = 0;
		driver.autoFill(array, column, excel);
	
		logger.info("End fill data to form create");
		
		for(int c = 1; c<=5; c++){
			if(!driver.getElenment("id=route_names_"+c).getAttribute("value").equals(""))
				routeCount+=1;
		}
		logger.info("Bam submit");
		driver.click("id=buttonSubmit", timeWaitJs);
	
		/*
		 * Kiểm tra xem có thông báo lỗi không. Nếu có in kết quả tạo đơn thất bại
		 * và chuyển sang đơn tiếp theo
		 */
		//Tìm tất cả các thông báo lỗi, nếu có 1 thông báo khác rỗng thì FALSE
		List<WebElement> ls = driver.getListElenment("className=invalid-msg");
		for(WebElement e : ls){
			if(!e.getText().contentEquals("")){
				logger.error("Ringi message error "+e.getText());
	    		return false;
	    	}
		}	
		/*
		 * Nếu không có thông báo lỗi thì click tiếp vào ô "OK", kết thúc nhập đơn và
		 * in kết quả đã tạo đơn thành công
		 */
		logger.info("bam OK");
		driver.click("id=insertButton", timeWaitAjax);
			
		int indexMessage = Common.getIndexFromArray(array, "Message");
		int indexLinkRedirect =  Common.getIndexFromArray(array, "urlRedirect");
		//Check GUI
		logger.info("Start check GUI");
		if(!excel.getStringData(column, indexMessage).equals(driver.getText("id=flashMessage"))
				|| !(linkStart+excel.getStringData(column, indexLinkRedirect)).equals(driver.getCurrentUrl())){
			logger.error("Error GUI");
			return false;
		}
		logger.info("End check GUI");
			        
		ResultSet rs1 = null;
		rs1 = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal ORDER BY id DESC LIMIT 1");
		rs1.next();
		excel.printStringIntoExcel(column, Common.getIndexFromArray(array, "newID"), rs1.getString("proposal_id"));
		excel.write();
		
		if(excel.getStringData(column, Common.getIndexFromArray(array, "chooseValidate")).equals("N")){
			excel.finish();
			return true;
		}else{
			int indexProposals = Common.getIndexFromArray(array, "id_proposal");
			int indexTsProposal = Common.getIndexFromArray(array, "id_ts_proposal");
			int indexTsProposalHistories = Common.getIndexFromArray(array, "id_ts_proposal_histories");
			int indexTsProposalRoute = Common.getIndexFromArray(array, "id_ts_proposal_route");
			
			//Check database
			logger.info("Start check DB");
			ResultSet rs;
			/*
			 * Kiểm tra bảng proposals
			 */
			//Truy vấn lấy proposals theo ID
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM proposals ORDER BY id DESC LIMIT 1");
			rs.next();
			/*
			 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!db.checkDataInTableProposals(excel, column, indexProposals, rs)){
				logger.error("Failse in table proposals");
				return false;
			}
			
			//Check db ts_proposal
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal ORDER BY id DESC LIMIT 1");
			rs.next();
			/*
			 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!db.checkDataInTableTsProposal(excel, column, indexTsProposal, rs)){
				logger.error("Failse in table ts_proposal");
				return false;
			}

			/*
			 * Kiểm tra bảng ts_proposal_histories
			 */
			//Truy vấn lấy ts_proposal_histories theo ID
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_histories ORDER BY id DESC LIMIT 1");
			rs.next();

			/*
			 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!db.checkDataInTableTsProposalHistories(excel, column, indexTsProposalHistories, rs)){
				logger.error("Failse in table ts_proposals_histories");
				return false;
			}
			
			/*
			 * Kiểm tra bảng ts_proposal_route
			 * Nếu số lượng route > 0 thì mới check
			 */
			 //Đếm số route
			if(routeCount > 0){
				//Truy vấn lấy ts_proposal_route theo ID, lấy số lượng bản ghi = số lượng route
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_route WHERE proposal_id = "+excel.getStringData(column, indexTsProposalRoute)+" ORDER BY id ASC LIMIT "+routeCount);
			
				/*
				 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Số lần lặp = số route, sau mỗi lần lặp cột chứa data đầu tiên + 5 ô
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				while(rs.next()){
					if(!db.checkDataInTableTsProposalRoute(excel, column, indexTsProposalRoute, rs)){		
						logger.error("Failse in table ts_proposal_route");
						return false;
					}
					indexTsProposalRoute += sumColumnRoute;
				}
			}
			logger.info("End check db");
			excel.finish();
			return true;
		}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto setting duyệt đơn đi lại
	 * @Author: hieuht
	 * @Date: 28/09/2016
	 */
	public boolean settingApprove(int column) throws Exception{
		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		//Khởi tạo logger
		logger=Logger.getLogger("settingApprove");

		//Tạo file output excel, truy cập sheet 2
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
								new ObjectInput("xpath=//*[@id='sortable_manage']/li/div/table/tbody/tr/td[2]/select", 2, 1, 0, timeWaitAjax),
								new ObjectInput("", 2, 4, 0, "use_group3", 0),
								new ObjectInput("id=comment", 2, 2, 0, 0),
								new ObjectInput("", 0, 0, 0, "Message", 0),
								new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
								new ObjectInput("", 0, 0, 0, "status_ts_proposal", 0),
								new ObjectInput("", 0, 0, 0, "id_ts_proposal_histories", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "creator_ts_proposal_process", 0)
			};
			if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
				login(excel, "Setting Approve", column, array);
			}			
			
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
			driver.autoFill(array, column, excel);
			logger.info("End fill data to form setting approve");
		
			//Bấm submit
			driver.click("id=buttonSubmit", timeWaitJs);
			if(!driver.getText("className=invalid-msg").equals("")){
				logger.error("Error message Ringi: " + driver.getText("className=invalid-msg"));
	    		return false;
			} else if(!driver.getText("id=msg_duplicate").equals("")) {
				logger.error("Error message Ringi: " + driver.getText("id=msg_duplicate"));
	    		return false;
			}
			//Bấm "OK" kết thúc setting duyệt
			driver.click("id=buttonSubmitSetting", 0);
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
				int indexTsProposalProcess = Common.getIndexFromArray(array, "creator_ts_proposal_process");
				int sumColumnProcess = 9;
				int indexTsProposal = Common.getIndexFromArray(array, "status_ts_proposal");
				int indexTsProposalHistories = Common.getIndexFromArray(array, "id_ts_proposal_histories");
				//Check DB			
				logger.info("Start check db");
				//Kiểm tra bảng ts_proposals
				//Truy vấn lấy ts_proposal theo ID
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal WHERE proposal_id = "+idProposal);
				rs.next();
				/*
				 * Kiểm tra dữ liệu trong excel với cột "status"
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				if(!(excel.getStringData(column, indexTsProposal).equals(DatabaseAction.getStringData(rs, "status")))){
					logger.error("Failse in table ts_proposal");
					return false;
				}
	
				//Kiểm tra bảng ts_proposal_histories
				//Truy vấn lấy ts_proposal_histories theo ID
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_histories order by id DESC LIMIT 1");
				rs.next();
				/*
				 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				if(!db.checkDataInTableTsProposalHistories(excel, column, indexTsProposalHistories, rs)){
					logger.error("Failse in table ts_proposal_histories");
					return false;
				}
				
				//Kiểm tra bảng ts_proposal_process
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_process WHERE proposal_id = "+idProposal+" ORDER BY id ASC LIMIT 5");
				while(rs.next()){
					if(!db.checkDataInTableTsProposalProcess(excel, column, indexTsProposalProcess, rs)){
						logger.error("Failse in table ts_proposal_process");
						return false;
					}
					indexTsProposalProcess += sumColumnProcess;
				}
				logger.info("End check db");
				return true;
		}
	}
		
	@Test
	/*
	 * Hàm thực hiện auto back đơn đi lại
	 * @Author: hieuht
	 * @Date: 17/10/2016
	 */
	public boolean back(int column) throws Exception{
		//Khởi tạo logger
		logger=Logger.getLogger("Back Proposal");

		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		logger.info("Read data form excel");
		//Tạo file output excel, truy cập sheet 6
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);  
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 31);
		excel.accessSheet("Back");
		
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, setActionChild, actionChild
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
								new ObjectInput("", 0, 0, 0, "numberUser", 0),
								new ObjectInput("", 0, 0, 0, "idProposal", 0),
								new ObjectInput("", 0, 0, 0, "choose", 0),
								new ObjectInput("id=content_proposal", 2, 2, 0, "comment", 0),
								new ObjectInput("", 0, 0, 0, "message", 0),
								new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
								new ObjectInput("", 0, 0, 0, "lastOperator_ts", 0),
								new ObjectInput("", 0, 0, 0, "status_ts", 0),
								new ObjectInput("", 0, 0, 0, "idTsProposalHistories", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "approval_status", 0),
								new ObjectInput("", 0, 0, 0, "status_all", 0),
								new ObjectInput("", 0, 0, 0, "comment_ts", 0),
				};
			if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
				login(excel, "Back", column, array);
			}	
		
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
			alert = driver.closeAlertAndGetItsText(timeWaitAjax*2);
			if(!alert.equals("差し戻します。よろしいですか？")){
				logger.error("Ringi alert message "+ alert);
				return false;
			}	
			if(!driver.getText("id=flashMessage").equals("差戻に成功しました。")){
				logger.error("Ringi redirect message "+ driver.getText("id=flashMessage"));
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
				int indexTsProposalProcess1 = Common.getIndexFromArray(array, "approval_status");
				int indexTsProposalProcess2 = Common.getIndexFromArray(array, "status_all");
				int indexTsProposalProcess3 = Common.getIndexFromArray(array, "comment_ts");
				int sumProcess = 3;		
				int indexLastOperatorTsProposal = Common.getIndexFromArray(array, "lastOperator_ts");
				int indexStatusTsProposal = Common.getIndexFromArray(array, "status_ts");
				int indexTsProposalHistories = Common.getIndexFromArray(array, "idTsProposalHistories");
				
				//Check db
				ResultSet rs;
				//Kiểm tra bảng ts_proposal_histories
				//Truy vấn lấy ts_proposal_histories theo ID
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_histories order by id DESC LIMIT 1");
				rs.next();
				/*
				 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				if(!db.checkDataInTableTsProposalHistories(excel, column, indexTsProposalHistories, rs)){
					logger.error("Failse in table ts_proposal_histories");
					return false;
				}
				
				//Kiểm tra bảng ts_proposals
				//Truy vấn lấy ts_proposal theo ID
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal WHERE proposal_id = "+idProposal);
				rs.next();
				/*
				 * Kiểm tra dữ liệu trong excel với cột "status"
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				if(!(excel.getStringData(column, indexStatusTsProposal).equals(DatabaseAction.getStringData(rs, "status")))
						|| !(excel.getStringData(column, indexLastOperatorTsProposal).equals(DatabaseAction.getStringData(rs, "last_operator")))){
					logger.error("Failse in table ts_proposal");
					return false;
				}
				
				//Kiểm tra bảng ts_proposal_process
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_process WHERE proposal_id = "+idProposal+" ORDER BY id ASC LIMIT 5");
				while(rs.next()){
					if(!(excel.getStringData(column, indexTsProposalProcess1).equals(DatabaseAction.getStringData(rs, "approval_status")))
						|| !(excel.getStringData(column, indexTsProposalProcess2).equals(DatabaseAction.getStringData(rs, "status_all")))
						|| !(excel.getStringData(column, indexTsProposalProcess3).equals(DatabaseAction.getStringData(rs, "comment")))){
						logger.error("Failse in table ts_proposal_process");
						return false;
					}
					indexTsProposalProcess1+= sumProcess;
					indexTsProposalProcess2+= sumProcess;
					indexTsProposalProcess3+= sumProcess;
				}
				logger.info("End check db");
				return true;
			}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto xác định đơn
	 * @Author: hieuht
	 * @Date: 20/10/2016
	 */
	public boolean determine(int column) throws Exception{
		//Khởi tạo logger
		logger=Logger.getLogger("Determine Proposal");

		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		logger.info("Read data form excel");
		
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild); 
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 32);
		excel.accessSheet("Determine");
		
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, setActionChild, actionChild
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
								new ObjectInput("", 0, 0, 0, "numberUser", 0),
								new ObjectInput("", 0, 0, 0, "idProposal", 0),
								new ObjectInput("id=input_comment", 2, 2, 0, "comment", 0),
								new ObjectInput("", 0, 0, 0, "message", 0),
								new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
								new ObjectInput("", 0, 0, 0, "status", 0),
								new ObjectInput("", 0, 0, 0, "determine_content", 0),
								new ObjectInput("", 0, 0, 0, "user_id_determine", 0),
								new ObjectInput("", 0, 0, 0, "idTsProposalHistories", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								};
			if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
				login(excel, "Determine", column, array);
			}	
				
			String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
			logger.info("Start fill data to form Determine proposal");
			driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
			driver.autoFill(array, column, excel);
			driver.click("id=submitDetermine", timeWaitJs);
			driver.closeAlertAndGetItsText(timeWaitAjax*2);
			
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
				int indexStatus = Common.getIndexFromArray(array, "status");
				int indexDetermineContent = Common.getIndexFromArray(array, "determine_content");
				int indexUserIdDetermine = Common.getIndexFromArray(array, "user_id_determine");
				int indexTsProposalHistories = Common.getIndexFromArray(array, "idTsProposalHistories");
				//Check db
				ResultSet rs;
				//Kiểm tra bảng ts_proposal_histories
				//Truy vấn lấy ts_proposal_histories theo ID
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_histories order by id DESC LIMIT 1");
				rs.next();
				/*
				 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				if(!db.checkDataInTableTsProposalHistories(excel, column, indexTsProposalHistories, rs)){
					logger.error("Failse in table ts_proposal_histories");
					return false;
				}
				
				//Kiểm tra bảng ts_proposals
				//Truy vấn lấy ts_proposal theo ID
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal WHERE proposal_id = "+idProposal);
				rs.next();
				/*
				 * Kiểm tra dữ liệu trong excel với cột "status"
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				if(!(excel.getStringData(column, indexStatus).equals(DatabaseAction.getStringData(rs, "status")))
						|| !(excel.getStringData(column, indexDetermineContent).equals(DatabaseAction.getStringData(rs, "determine_content")))
						|| !(excel.getStringData(column, indexUserIdDetermine).equals(DatabaseAction.getStringData(rs, "user_id_determine")))){
					logger.error("Failse in table ts_proposal");
					return false;
				}
				logger.info("End check db");
				return true;
			}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto reject đơn đi lại
	 * @Author: hieuht
	 * @Date: 17/10/2016
	 */
	public boolean reject(int column) throws Exception{
		//Khởi tạo logger
		logger=Logger.getLogger("Reject Proposal");
	
		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		logger.info("Read data form excel");
		//Tạo file output excel, truy cập sheet 7
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild); 
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 31);
		excel.accessSheet("Reject");
		
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, setActionChild, actionChild
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
				new ObjectInput("", 0, 0, 0, "numberUser", 0),
				new ObjectInput("", 0, 0, 0, "idProposal", 0),
				new ObjectInput("", 0, 0, 0, "choose", 0),
				new ObjectInput("id=content_proposal", 2, 2, 0, "comment", 0),
				new ObjectInput("", 0, 0, 0, "message", 0),
				new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
				new ObjectInput("", 0, 0, 0, "lastOperator_ts", 0),
				new ObjectInput("", 0, 0, 0, "status_ts", 0),
				new ObjectInput("", 0, 0, 0, "idTsProposalHistories", 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, "approval_status", 0),
				new ObjectInput("", 0, 0, 0, "status_all", 0),
				new ObjectInput("", 0, 0, 0, "comment_ts", 0),
				};
		
			if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
				login(excel, "Reject", column, array);
			}	
		
			String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
			String alert = "";	
			logger.info("Start fill data to form reject proposal");
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
			driver.click("id=reject_proposal", timeWaitJs);
			logger.info("End fill form");
			alert = driver.closeAlertAndGetItsText(timeWaitAjax*2);
			if(!alert.equals("却下します。よろしいですか？")){
				logger.error("Ringi alert error "+alert);
				return false;
			}	
			if(!driver.getText("id=flashMessage").equals("却下に成功しました。")){
				logger.error("Ringi redirect message error "+driver.getText("id=flashMessage"));
				return false;
			}	
			logger.info("Reject don thanh cong");
			
			int indexMessage = Common.getIndexFromArray(array, "message");
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
			}else{
				int indexTsProposalProcess1 = Common.getIndexFromArray(array, "approval_status");
				int indexTsProposalProcess2 = Common.getIndexFromArray(array, "status_all");
				int indexTsProposalProcess3 = Common.getIndexFromArray(array, "comment_ts");
				int sumProcess = 3;		
				int indexLastOperatorTsProposal = Common.getIndexFromArray(array, "lastOperator_ts");
				int indexStatusTsProposal = Common.getIndexFromArray(array, "status_ts");
				int indexTsProposalHistories = Common.getIndexFromArray(array, "idTsProposalHistories");
				
				//Check db
				ResultSet rs;
				//Kiểm tra bảng ts_proposal_histories
				//Truy vấn lấy ts_proposal_histories theo ID
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_histories order by id DESC LIMIT 1");
				rs.next();
				/*
				 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				if(!db.checkDataInTableTsProposalHistories(excel, column, indexTsProposalHistories, rs)){
					logger.error("Failse in table ts_proposal_histories");
					return false;
				}
				
				//Kiểm tra bảng ts_proposals
				//Truy vấn lấy ts_proposal theo ID
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal WHERE proposal_id = "+idProposal);
				rs.next();
				/*
				 * Kiểm tra dữ liệu trong excel với cột "status"
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				if(!(excel.getStringData(column, indexStatusTsProposal).equals(DatabaseAction.getStringData(rs, "status")))
						|| !(excel.getStringData(column, indexLastOperatorTsProposal).equals(DatabaseAction.getStringData(rs, "last_operator")))){
					logger.error("Failse in table ts_proposal");
					return false;
				}
				
				//Kiểm tra bảng ts_proposal_process
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_process WHERE proposal_id = "+idProposal+" ORDER BY id ASC LIMIT 5");
				while(rs.next()){
					if(!(excel.getStringData(column, indexTsProposalProcess1).equals(DatabaseAction.getStringData(rs, "approval_status")))
						|| !(excel.getStringData(column, indexTsProposalProcess2).equals(DatabaseAction.getStringData(rs, "status_all")))
						|| !(excel.getStringData(column, indexTsProposalProcess3).equals(DatabaseAction.getStringData(rs, "comment")))){
						logger.error("Failse in table ts_proposal_process");
						return false;
					}
					indexTsProposalProcess1+= sumProcess;
					indexTsProposalProcess2+= sumProcess;
					indexTsProposalProcess3+= sumProcess;
				}
				logger.info("End check db");
				return true;
			}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto đề xuất lại đơn đi lại
	 * @Author: hieuht
	 * @Date: 16/09/2016
	 */
	public boolean createProposalAgain(int column) throws Exception {
		//Khởi tạo logger
		logger=Logger.getLogger("MultiProposalAgain");
		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		logger.info("Read data form excel");
		//Tạo file output excel, truy cập sheet 3
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);  
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 30);
		excel.accessSheet("Create Proposal Again");
		
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, setActionChild, actionChild
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
								new ObjectInput("", 0, 0, 0, "numberUser", 0),
								new ObjectInput("", 0, 0, 0, "idProposal", 0),
								new ObjectInput("id=UserCreateProposal", 0, 0, 0, "userHelp", 0),
								new ObjectInput("", 0, 0, 0, "idUserHelp", 0),
								new ObjectInput("id=company_code", 2, 1, 0, timeWaitAjax),
								new ObjectInput("id=group_code", 2, 1, 0, timeWaitAjax),
								new ObjectInput("id=division_code", 2, 1, 0, timeWaitAjax),
								new ObjectInput("id=type_id", 2, 1, 0, timeWaitAjax),
								new ObjectInput("id=category_id", 2, 1, 0, timeWaitAjax),
								new ObjectInput("name=data[User][change_date]", 2, 3, 0, timeWaitJs),
								new ObjectInput("", 0, 4, 0, "supplyCost", timeWaitJs),
								new ObjectInput("id=reason_type", 2, 1, 0, 0),
								new ObjectInput("id=post_code", 2, 2, 0, timeWaitAjax),
								new ObjectInput("id=address", 2, 2, 0, 0),
								new ObjectInput("id=route_names_1", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-1", 2, 2, 0, timeWaitAjax),
								new ObjectInput("id=ts-distance-to-1", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_1", 2, 2, 0, 0),
								new ObjectInput("id=route_names_2", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-2", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-2", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_2", 2, 2, 0, 0),
								new ObjectInput("id=route_names_3", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-3", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-3", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_3", 2, 2, 0, 0),
								new ObjectInput("id=route_names_4", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-4", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-4", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_4", 2, 2, 0, 0),
								new ObjectInput("id=route_names_5", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-5", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-5", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_5", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_one_day", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_one_month", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_three_month", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_six_month", 2, 2, 0, 0),
								new ObjectInput("id=comment", 2, 2, 0, 0),
								new ObjectInput("id=origin-input", 2, 2, 0, 0),
								new ObjectInput("id=destination-input", 2, 2, 0, timeWaitAjax),
								new ObjectInput("", 0, 0, 0, "Message", 0),
								new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
								new ObjectInput("", 0, 0, 0, "id_proposal", 0),
								new ObjectInput("", 0, 0, 0, "start_ts_proposal", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "id_ts_proposal_histories", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "id_ts_proposal_route", 0),
			};	
			
			if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
				login(excel, "Create Proposal Again", column, array);
			}	
		
			int sumColumnRoute = 5;
			int supplyCost = Common.getIndexFromArray(array, "supplyCost");
			if(excel.getStringData(column, supplyCost).equals("1")){
				array[supplyCost].setId("id=supply_cost_yes");
				array[supplyCost].setActon(1);
			}else if(excel.getStringData(column, supplyCost).equals("0")){
				array[supplyCost].setId("id=supply_cost_no");
				array[supplyCost].setActon(1);
			}
			String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
			driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
			//Chọn vào button đề xuất lại

			driver.click("className=proposal_again", timeWaitAjax);
			logger.info("Start fill data to form de xuat lai");
			
			int helpBy = Common.getIndexFromArray(array, "userHelp");
			int idHelpBy = Common.getIndexFromArray(array, "idUserHelp");
			if(!excel.getStringData(column, helpBy).equals("")){
				driver.sendKey("id=UserCreateProposal", excel.getStringData(column, helpBy), timeWaitAjax*2);
				try {
					ResultSet rs = null;
					rs = DatabaseAction.stmt.executeQuery("SELECT id FROM users where username = "+excel.getStringData(column, idHelpBy));
					rs.next();
					driver.click("xpath=//li[@class='"+rs.getString("id")+"']", timeWaitAjax);
				} catch (Exception e) {
					logger.error("Khong tim thay user tao thay");
					return false;
				}
			}
			
			driver.autoFill(array, column, excel);
			logger.info("End fill data to form de xuat lai");
			
			int routeCount = 0;
			for(int c = 1; c<=5; c++){
				if(!driver.getElenment("id=route_names_"+c).getAttribute("value").equals(""))
					routeCount+=1;
			}
			logger.info("Bam submit");
			driver.click("id=buttonSubmit", timeWaitJs);
			/*
			 * Kiểm tra xem có thông báo lỗi không. Nếu có in kết quả tạo đơn thất bại
			 * và chuyển sang đơn tiếp theo
			 */		
			//Tìm tất cả các thông báo lỗi, nếu có 1 thông báo khác rỗng thì FALSE
			List<WebElement> ls = driver.getListElenment("className=invalid-msg");
			for(WebElement e : ls){
				if(!e.getText().contentEquals("")){
		    		logger.error("Ringi error message "+e.getText());
		    		return false;
		    	}
			}	
			/*
			 * Nếu không có thông báo lỗi thì click tiếp vào ô "OK", kết thúc nhập đơn và
			 * in kết quả đã tạo đơn thành công
			 */
			logger.info("bam OK");
			driver.click("id=insertButton", timeWaitAjax);
			
			int indexMessage = Common.getIndexFromArray(array, "Message");
			int indexLinkRedirect =  Common.getIndexFromArray(array, "urlRedirect");
			//Check GUI
			logger.info("Start check GUI");
			if(!excel.getStringData(column, indexMessage).equals(driver.getText("id=flashMessage"))
					|| !(linkStart+excel.getStringData(column, indexLinkRedirect)).equals(driver.getCurrentUrl())){
				logger.error("Error GUI");
				return false;
			}
			logger.info("End check GUI");
			
			if(excel.getStringData(column, Common.getIndexFromArray(array, "chooseValidate")).equals("N")){
				return true;
			}else{
				int indexIdProposal = Common.getIndexFromArray(array, "idProposal");
				
				int indexIdProposalCheck = Common.getIndexFromArray(array, "id_proposal");
				int indexTsProposalHistories = Common.getIndexFromArray(array, "id_ts_proposal_histories");
				int indexTsProposalRoute = Common.getIndexFromArray(array, "id_ts_proposal_route");
				int indexStartTsProposal = Common.getIndexFromArray(array, "start_ts_proposal");
				
				//Check database
				logger.info("Start check DB");
				ResultSet rs;
				
				//Kiểm tra bảng proposals
				//Truy vấn lấy proposals theo ID
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM proposals WHERE id = "+idProposal);
				rs.next();
				/*
				 * Kiểm tra dữ liệu trong excel với cột "helped_by"
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				if(!(excel.getStringData(column, indexIdProposalCheck).equals(DatabaseAction.getStringData(rs, "helped_by")))){
					return false;
				}
				//Check column help_by db ts_proposal
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal WHERE proposal_id = "+idProposal);
				rs.next();
				/*
				 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				ObjectInput[] arrayTs = {	new ObjectInput("helped_by", 2),
											new ObjectInput("company_id", 2),
											new ObjectInput("group_id", 2),
											new ObjectInput("division_id", 2),
											new ObjectInput("user_type_id", 2),
											new ObjectInput("category", 2),
											new ObjectInput("change_date", 3),
											new ObjectInput("supply_cost", 2),
											new ObjectInput("reason_type", 2),
											new ObjectInput("post_code", 2),
											new ObjectInput("address", 2),
											new ObjectInput("comment", 2),
											new ObjectInput("status", 2),
											new ObjectInput("re_proposal", 2),
											new ObjectInput("total_cost_one_day", 2),
											new ObjectInput("total_cost_one_month", 2),
											new ObjectInput("total_cost_three_month", 2),
											new ObjectInput("total_cost_six_month", 2),
											new ObjectInput("start_point", 2),
											new ObjectInput("end_point", 2),
											new ObjectInput("old_proposal_id", 2)};
				if(!DatabaseAction.compare(arrayTs, logger, excel, column, indexStartTsProposal, rs)){
					logger.error("Failse in table ts_proposal");
					return false;
				}
				/*
				 * Kiểm tra bảng ts_proposal_histories
				 */
				//Truy vấn lấy ts_proposal_histories theo ID
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_histories ORDER BY id DESC LIMIT 1");
				rs.next();
	
				/*
				 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				if(!db.checkDataInTableTsProposalHistories(excel, column, indexTsProposalHistories, rs)){
					logger.error("Failse in table ts_proposals_histories");
					return false;
				}
				/*
				 * Kiểm tra bảng ts_proposal_route
				 * Nếu số lượng route > 0 thì mới check
				 */
				 //Đếm số route
				if(routeCount > 0){
					//Truy vấn lấy ts_proposal_route theo ID, lấy số lượng bản ghi = số lượng route
					rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_route WHERE proposal_id = "+indexIdProposal+" ORDER BY id ASC LIMIT "+routeCount);
				
					/*
					 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
					 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
					 * Số lần lặp = số route, sau mỗi lần lặp cột chứa data đầu tiên + 5 ô
					 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
					 */
					while(rs.next()){
						if(!db.checkDataInTableTsProposalRoute(excel, column, indexTsProposalRoute, rs)){
							logger.error("Failse in table ts_proposal_route");
							return false;
						}
						indexTsProposalRoute += sumColumnRoute;
					}
				}
				logger.info("End check db");
				return true;
			}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto setting back đơn đi lại
	 * @Author: hieuht
	 * @Date: 13/10/2016
	 */
	public boolean settingBack(int column) throws Exception{
		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		//Khởi tạo logger
		logger=Logger.getLogger("SettingBackProposal");
		//Tạo file output excel, truy cập sheet 2
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);   
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 29);
		excel.accessSheet("Setting Approve_Back");
		
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
								new ObjectInput("", 0, 0, 0, "numberUser", 0),
								new ObjectInput("", 0, 0, 0, "idProposal", 0),
								new ObjectInput("id=comment", 2, 2, 0, 0),
								new ObjectInput("", 0, 0, 0, "Message", 0),
								new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
								new ObjectInput("", 0, 0, 0, "status_ts_proposal", 0),
								new ObjectInput("", 0, 0, 0, "id_ts_proposal_histories", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
		};
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
			login(excel, "Setting Approve_Back", column, array);
		}	
			
		String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
		link.replace("[%IDPROPOSAL]", idProposal);
		driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
		logger.info("Start fill data to form setting approve");
		driver.autoFill(array, column, excel);
		logger.info("End fill data to form setting approve");
		
		//Bấm submit
		try {
			driver.click("id=buttonBack", timeWaitJs);
			driver.closeAlertAndGetItsText(timeWaitAjax*2);
		} catch (Exception e) {
			logger.error("Chua nhap dung yeu cau");
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
			int indexStatusTsProposal = Common.getIndexFromArray(array, "status_ts_proposal");
			int indexTsProposalHistories = Common.getIndexFromArray(array, "id_ts_proposal_histories");		
			//Check DB
			ResultSet rs;
			logger.info("Start check db");
			//Kiểm tra bảng ts_proposals
			//Truy vấn lấy ts_proposal theo ID
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal WHERE proposal_id = "+idProposal);
			rs.next();
			/*
			 * Kiểm tra dữ liệu trong excel với cột "status"
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!(excel.getStringData(column, indexStatusTsProposal).equals(DatabaseAction.getStringData(rs, "status")))){
				logger.error("Failse in table ts_proposal");
				return false;
			}
			
			//Kiểm tra bảng ts_proposal_histories
			//Truy vấn lấy ts_proposal_histories theo ID
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_histories order by id DESC LIMIT 1");
			rs.next();
			/*
			 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!db.checkDataInTableTsProposalHistories(excel, column, indexTsProposalHistories, rs)){
				logger.error("Failse in table ts_proposal_histories");
				return false;
			}
			logger.info("End check db");
			return true;
		}
	}

	@Test
	/*
	 * Hàm thực hiện auto chỉnh sửa đơn
	 * @Author: hieuht
	 * @Date: 21/10/2016
	 */
	public boolean edit(int column) throws Exception{
		//Khởi tạo logger
		logger=Logger.getLogger("editProposal");	
		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		logger.info("Read data form excel");
		//Tạo file output excel, truy cập sheet 1
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild); 
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 33);
		excel.accessSheet("Edit Proposal");
		
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, setActionChild, actionChild
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
								new ObjectInput("", 0, 0, 0, "numberUser", 0),
								new ObjectInput("", 0, 0, 0, "idProposal", 0),
								new ObjectInput("name=data[User][change_date]", 2, 3, 0, 0),
								new ObjectInput("id=post_code", 2, 2, 0, timeWaitAjax),
								new ObjectInput("id=address", 2, 2, 0, 0),
								new ObjectInput("id=route_names_1", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-1", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-1", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_1", 2, 2, 0, 0),
								new ObjectInput("id=route_names_2", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-2", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-2", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_2", 2, 2, 0, 0),
								new ObjectInput("id=route_names_3", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-3", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-3", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_3", 2, 2, 0, 0),
								new ObjectInput("id=route_names_4", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-4", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-4", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_4", 2, 2, 0, 0),
								new ObjectInput("id=route_names_5", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-from-5", 2, 2, 0, 0),
								new ObjectInput("id=ts-distance-to-5", 2, 2, 0, 0),
								new ObjectInput("id=cost_of_route_one_day_5", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_one_day", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_one_month", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_three_month", 2, 2, 0, 0),
								new ObjectInput("id=total_cost_six_month", 2, 2, 0, 0),
								new ObjectInput("id=comment", 2, 2, 0, 0),
								new ObjectInput("id=origin-input", 2, 2, 0, 0),
								new ObjectInput("id=destination-input", 2, 2, 0, timeWaitAjax),
								new ObjectInput("", 0, 0, 0, "Message", 0),
								new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
								new ObjectInput("", 0, 0, 0, "id_proposal", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "id_ts_proposal", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0), 
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "id_ts_proposal_histories", 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, 0),
								new ObjectInput("", 0, 0, 0, "id_ts_proposal_route", 0),
			};	
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
			login(excel, "Edit Proposal", column, array);
		}	
		
		int sumColumnRoute = 5;
		String idProposal = excel.getFormulaCellData(column, Common.getIndexFromArray(array, "idProposal"));
		driver.open(linkStart+link.replace("[%IDPROPOSAL]", idProposal));
		driver.click("className=ts_edit_proposal", timeWaitAjax);
		
		logger.info("Start fill data to form edit");
		driver.autoFill(array, column, excel);
		logger.info("End fill data to form edit");
		
		int routeCount = 0;
		for(int c = 1; c<=5; c++){
			if(!driver.getElenment("id=route_names_"+c).getAttribute("value").equals(""))
				routeCount+=1;
		}
		logger.info("Bam submit");
		driver.click("id=buttonSubmit", timeWaitJs);
		/*
		 * Kiểm tra xem có thông báo lỗi không. Nếu có in kết quả tạo đơn thất bại
		 * và chuyển sang đơn tiếp theo
		 */
		//Tìm tất cả các thông báo lỗi, nếu có 1 thông báo khác rỗng thì FALSE
		List<WebElement> ls = driver.getListElenment("className=invalid-msg");
		for(WebElement e : ls){
			if(!e.getText().contentEquals("")){
	    		logger.error("Ringi error message "+e.getText());
	    		return false;
	    	}
		}		
		/*
		 * Nếu không có thông báo lỗi thì click tiếp vào ô "OK", kết thúc nhập đơn và
		 * in kết quả đã tạo đơn thành công
		 */
		logger.info("bam OK");
		driver.click("id=insertButton", timeWaitAjax);
		
		int indexMessage = Common.getIndexFromArray(array, "Message");
		int indexLinkRedirect =  Common.getIndexFromArray(array, "urlRedirect");
		//Check GUI
		logger.info("Start check GUI");
		if(!excel.getStringData(column, indexMessage).equals(driver.getText("id=flashMessage"))
				|| !(linkStart+excel.getStringData(column, indexLinkRedirect)).equals(driver.getCurrentUrl())){
			logger.error("Error GUI");
			return false;
		}
		logger.info("End check GUI");
		
		if(excel.getStringData(column, Common.getIndexFromArray(array, "chooseValidate")).equals("N")){
			return true;
		}else{
			int indexProposals = Common.getIndexFromArray(array, "id_proposal");
			int indexTsProposal = Common.getIndexFromArray(array, "id_ts_proposal");
			int indexTsProposalHistories = Common.getIndexFromArray(array, "id_ts_proposal_histories");
			int indexTsProposalRoute = Common.getIndexFromArray(array, "id_ts_proposal_route");
			
			//Check database
			logger.info("Start check DB");
			ResultSet rs;
			/*
			 * Kiểm tra bảng proposals
			 */
			//Truy vấn lấy proposals theo ID
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM proposals ORDER BY id DESC LIMIT 1");
			rs.next();
			/*
			 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!db.checkDataInTableProposals(excel, column, indexProposals, rs)){
				logger.error("Failse in table proposals");
				return false;
			}
			
			//Check db ts_proposal
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal ORDER BY id DESC LIMIT 1");
			rs.next();
			/*
			 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!db.checkDataInTableTsProposal(excel, column, indexTsProposal, rs)){
				logger.error("Failse in table ts_proposal");
				return false;
			}
			/*
			 * Kiểm tra bảng ts_proposal_histories
			 */
			//Truy vấn lấy ts_proposal_histories theo ID
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_histories ORDER BY id DESC LIMIT 1");
			rs.next();

			/*
			 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!db.checkDataInTableTsProposalHistories(excel, column, indexTsProposalHistories, rs)){
				logger.error("Failse in table ts_proposals_histories");
				return false;
			}
			
			/*
			 * Kiểm tra bảng ts_proposal_route
			 * Nếu số lượng route > 0 thì mới check
			 */
			 //Đếm số route
			if(routeCount > 0){
				//Truy vấn lấy ts_proposal_route theo ID, lấy số lượng bản ghi = số lượng route
				rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_route WHERE proposal_id = "+excel.getStringData(column, indexTsProposalRoute)+" ORDER BY id ASC LIMIT "+routeCount);
			
				/*
				 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
				 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
				 * Số lần lặp = số route, sau mỗi lần lặp cột chứa data đầu tiên + 5 ô
				 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
				 */
				while(rs.next()){
					if(!db.checkDataInTableTsProposalRoute(excel, column, indexTsProposalRoute, rs)){
						logger.error("Failse in table ts_proposal_route");
						return false;
					}
					indexTsProposalRoute += sumColumnRoute;
				}
			}
			logger.info("End check db");
			return true;
		}
	}
		
	@Test
	/*
	 * Hàm thực hiện auto duyệt đơn
	 * @Author: hieuht
	 * @Date: 16/09/2016
	 */
	public boolean approve(int column) throws Exception {
		//Khởi tạo logger
		logger=Logger.getLogger("ApproveProposal");

		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		
		logger.info("Read data form excel");
		//Tạo file output excel, truy cập sheet 3
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 31);
		
		excel.accessSheet("Approve");

		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, setActionChild, actionChild
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
				new ObjectInput("", 0, 0, 0, "lastOperator_ts", 0),
				new ObjectInput("", 0, 0, 0, "status_ts", 0),
				new ObjectInput("", 0, 0, 0, "idTsProposalHistories", 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, "approval_status", 0),
				new ObjectInput("", 0, 0, 0, "status_all", 0),
				new ObjectInput("", 0, 0, 0, "comment_ts", 0),
				};
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
			login(excel, "Approve", column, array);
		}	
		
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
				driver.getElementFormList("id=chkSelected", 0).click();;
			}
			else if(excel.getStringData(column, indexChoose).equals("2")){
				driver.getElementFormList("id=chkSelected", 1).click();
			}
			driver.sendKey("id=content_proposal", excel.getStringData(column, Common.getIndexFromArray(array, "comment")), 0);
			driver.click("id=approved_proposal", timeWaitJs);
			alert = driver.closeAlertAndGetItsText(timeWaitAjax*2);
			if(!alert.equals("承認します。よろしいですか？")){
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
				alert = driver.closeAlertAndGetItsText(timeWaitAjax*2);
				if(!alert.equals("承認します。よろしいですか？")){
					logger.error("Ringi alert message "+alert);
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
			driver.click("id=approved_proposal", timeWaitJs);
			alert = driver.closeAlertAndGetItsText(timeWaitAjax*2);
			if(!alert.equals("承認します。よろしいですか？")){
				logger.error("Ringi alert message "+alert);
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
		    alert = driver.closeAlertAndGetItsText(timeWaitAjax*2);
			if(!alert.equals("承認します。よろしいですか？")){
				logger.error("Ringi alert message "+alert);
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
			int indexTsProposalProcess1 = Common.getIndexFromArray(array, "approval_status");
			int indexTsProposalProcess2 = Common.getIndexFromArray(array, "status_all");
			int indexTsProposalProcess3 = Common.getIndexFromArray(array, "comment_ts");
			int sumProcess = 3;
			int indexLastOperatorTsProposal = Common.getIndexFromArray(array, "lastOperator_ts");
			int indexStatusTsProposal = Common.getIndexFromArray(array, "status_ts");
			int indexTsProposalHistories = Common.getIndexFromArray(array, "idTsProposalHistories");
			
			//Check db
			ResultSet rs;
			//Kiểm tra bảng ts_proposal_histories
			//Truy vấn lấy ts_proposal_histories theo ID
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_histories order by id DESC LIMIT 1");
			rs.next();
			/*
			 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!db.checkDataInTableTsProposalHistories(excel, column, indexTsProposalHistories, rs)){
				logger.error("Failse in table ts_proposal_histories");
				return false;
			}
			
			//Kiểm tra bảng ts_proposals
			//Truy vấn lấy ts_proposal theo ID
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal WHERE proposal_id = "+idProposal);
			rs.next();
			/*
			 * Kiểm tra dữ liệu trong excel với cột "status"
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!(excel.getStringData(column, indexStatusTsProposal).equals(DatabaseAction.getStringData(rs, "status")))
					|| !(excel.getStringData(column, indexLastOperatorTsProposal).equals(DatabaseAction.getStringData(rs, "last_operator")))){
				logger.error("Failse in table ts_proposal");
				return false;
			}
			
			//Kiểm tra bảng ts_proposal_process
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_process WHERE proposal_id = "+idProposal+" ORDER BY id ASC LIMIT 5");
			while(rs.next()){
				if(!(excel.getStringData(column, indexTsProposalProcess1).equals(DatabaseAction.getStringData(rs, "approval_status")))
					|| !(excel.getStringData(column, indexTsProposalProcess2).equals(DatabaseAction.getStringData(rs, "status_all")))
					|| !(excel.getStringData(column, indexTsProposalProcess3).equals(DatabaseAction.getStringData(rs, "comment")))){
					logger.error("Failse in table ts_proposal_process");
					return false;
				}
				indexTsProposalProcess1+= sumProcess;
				indexTsProposalProcess2+= sumProcess;
				indexTsProposalProcess3+= sumProcess;
			}
			logger.info("End check db");
			excel.finish();
			return true;
		}
	}
	
	@Test
	/*
	 * Hàm thực hiện auto cancel đơn đi lại
	 * @Author: hieuht
	 * @Date: 20/10/2016
	 */
	public boolean cancel(int column) throws Exception{
		//Khởi tạo logger
		logger=Logger.getLogger("Cancel Proposal");

		//Thực hiện reset data nếu dữ liệu trong excel = 2
		if(resetData == 2){
			DatabaseAction.resetData(dbNameDelete, dbNameImport);
		}
		logger.info("Read data form excel");
		//Tạo file output excel, truy cập sheet 6
		ExcelAction excel = new ExcelAction(File, buildId, timeBuild);  
		excel.accessSheet("Config");
		String link = excel.getStringData(1, 31);
		excel.accessSheet("Cancel Approve");
		
		//0 là không xét, 1 là có xét, 2 là kiểm tra null rồi xét
		//ID, action, type, setActionChild, actionChild
		ObjectInput[] array = {	new ObjectInput("", 0, 0, 0, "chooseValidate", 0),
				new ObjectInput("", 0, 0, 0, "numberUser", 0),
				new ObjectInput("", 0, 0, 0, "idProposal", 0),
				new ObjectInput("", 0, 0, 0, "choose", 0),
				new ObjectInput("id=content_proposal", 1, 2, 0, "comment", 0),
				new ObjectInput("", 0, 0, 0, "message", 0),
				new ObjectInput("", 0, 0, 0, "urlRedirect", 0),
				new ObjectInput("", 0, 0, 0, "lastOperator_ts", 0),
				new ObjectInput("", 0, 0, 0, "idTsProposalHistories", 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, 0),
				new ObjectInput("", 0, 0, 0, "approval_status", 0),
				new ObjectInput("", 0, 0, 0, "comment_ts", 0),
				};
		if(!excel.getStringData(column, Common.getIndexFromArray(array, "numberUser")).equals("")){
			login(excel, "Cancel Approve", column, array);
		}	
		
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
		alert = driver.closeAlertAndGetItsText(timeWaitAjax*2);
		if(!alert.equals("承認取消しします。よろしいですか？")){
			logger.error("Ringi alert message "+alert);
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
			int indexTsProposalProcess1 = Common.getIndexFromArray(array, "approval_status");
			int indexTsProposalProcess3 = Common.getIndexFromArray(array, "comment_ts");
			int sumProcess = 2;			
			int indexLastOperatorTsProposal = Common.getIndexFromArray(array, "lastOperator_ts");
			int indexTsProposalHistories = Common.getIndexFromArray(array, "idTsProposalHistories");
			
			//Check db
			ResultSet rs;
			//Kiểm tra bảng ts_proposal_histories
			//Truy vấn lấy ts_proposal_histories theo ID
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_histories order by id DESC LIMIT 1");
			rs.next();
			/*
			 * Kiểm tra dữ liệu trong file excel với dữ liệu trong db truy vấn được
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!db.checkDataInTableTsProposalHistories(excel, column, indexTsProposalHistories, rs)){
				logger.error("Failse in table ts_proposal_histories");
				return false;
			}
			
			//Kiểm tra bảng ts_proposals
			//Truy vấn lấy ts_proposal theo ID
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal WHERE proposal_id = "+idProposal);
			rs.next();
			/*
			 * Kiểm tra dữ liệu trong excel với cột "status"
			 * Parameter: file excel chứa dữ liệu để check, dòng trong file excel, kết quả truy vấn, cột chứa data đầu tiên
			 * Nếu sai thì in kết quả "FALSE" vào excel và sang đơn tiếp theo, nếu đúng thì tiếp tục
			 */
			if(!(excel.getStringData(column, indexLastOperatorTsProposal).equals(DatabaseAction.getStringData(rs, "last_operator")))){
				logger.error("Failse in table ts_proposal");
				return false;
			}
			
			//Kiểm tra bảng ts_proposal_process
			rs = DatabaseAction.stmt.executeQuery("SELECT * FROM ts_proposal_process WHERE proposal_id = "+idProposal+" ORDER BY id ASC LIMIT 5");
			while(rs.next()){
				if(!(excel.getStringData(column, indexTsProposalProcess1).equals(DatabaseAction.getStringData(rs, "approval_status")))
					|| !(excel.getStringData(column, indexTsProposalProcess3).equals(DatabaseAction.getStringData(rs, "comment")))){
					logger.error("Failse in table ts_proposal_process");
					return false;
				}
				indexTsProposalProcess1+= sumProcess;
				indexTsProposalProcess3+= sumProcess;
			}
			logger.info("End check db");
			return true;
		}
	}
	
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
	 * Hàm thực hiện đóng trình duyệt FireFox và mở file output excel
	 * @Author: hieuht
	 * @Date:16/09/2016
	 */
	public void closeBrowser() throws IOException{
		//Khởi tạo logger
		logger=Logger.getLogger("closeBrower");
		driver.close();
		logger.info("Close firefox");
//		Runtime.getRuntime().exec("cmd /c start "+ pathFile);
		logger.info("Open output file successfully");
		endTime = System.currentTimeMillis();
		long totalTime = endTime - startTime;
		logger.info("Tong thoi gian chay: "+ totalTime/1000 + "s ~ "+ totalTime/1000/60+"p"+totalTime/1000%60+"s");
		logger.info("Tong time nghi: "+driver.getCountTime()/1000 + "s ~ "+driver.getCountTime()/1000/60+"p"+driver.getCountTime()/1000%60+"s");
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		Date date = new Date();
		logger.info("End autotest don di lai at " + dateFormat.format(date));
	}
}
