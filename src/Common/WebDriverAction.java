/*
 * Lớp thực hiện các tác động tới website
 * @Author: hieuht
 * @Date: 16/09/2016
 */

package Common;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.DateUtil;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

public class WebDriverAction {
	public WebDriver driver;
	public int row;
	public int countTime;
	//Mở trình duyệt firefox version 46
	public WebDriverAction() {
		this.driver = new FirefoxDriver();
	}
	
	/*
	 * Author: hieuht
	 * Trả về dữ liệu kiểu By dựa vào định danh truyền vào
	 * Parameter: String locator: định danh xác định đối tượng
	 * Output: By: Giá trị kiểu By
	 */
	public By getBy(String locator){
		By by=null;
		if(locator.startsWith("id=")){
			locator = locator.substring(3);
			by = By.id(locator);
		} else if(locator.startsWith("css=")){
			locator = locator.substring(4);
			by = By.cssSelector(locator);
		} else if(locator.startsWith("xpath=")){
			locator = locator.substring(6);
			by = By.xpath(locator);
		} else if(locator.startsWith("className=")){
			locator = locator.substring(10);
			by = By.className(locator);
		} else if(locator.startsWith("name=")){
			locator = locator.substring(5);
			by = By.name(locator);
		} else if(locator.startsWith("linkText=")){
			locator = locator.substring(9);
			by = By.linkText(locator);
		}
		return by;
	}
	
	/*
	 * Author: hieuht
	 * Trả về dữ liệu kiểu WebElement dựa vào định danh truyền vào
	 * Parameter: String locator: định danh xác định đối tượng
	 * Output: WebElement: Giá trị kiểu WebElement
	 */
	public WebElement getElenment(String locator){
		return driver.findElement(getBy(locator));
	}
	
	/*
	 * Author: hieuht
	 * Trả về list dữ liệu kiểu WebElement dựa vào định danh truyền vào
	 * Parameter: String locator: định danh xác định đối tượng
	 * Output: List<Webelement>: Giá trị kiểu List<Webelement>
	 */
	public List<WebElement> getListElenment(String locator){
		List<WebElement> ls = null;
		ls = driver.findElements(getBy(locator));
		return ls;
	}

	/*
	 * Author: hieuht
	 * Lấy 1 giá trị kiểu WebElement trong 1 list WebElement theo chỉ số index
	 * Parameter: String locator: định danh xác định đối tượng, 
	 * 			  int index: chỉ số cần lấy
	 * Output: Giá trị kiểu WebElement
	 */
	public WebElement getElementFormList(String locator, int index) throws Exception{
		return this.getListElenment(locator).get(index);
	}
	
	/*
	 * Author: hieuht
	 * Click vào 1 đối tượng dựa vào định danh
	 * Parameter: String locator: định danh xác định đối tượng
	 */
	public void click(String locator, int timeSleep) throws Exception{
		getElenment(locator).click();
		sleep(timeSleep);
	}
	
	/*
	 * Author: hieuht
	 * Trả về thuộc tính driver
	 * Output: WebDriver: dữ liệu nhận được
	 */
	public WebDriver getDriver(){
		return this.driver;
	}
	
	/*
	 * Author: hieuht
	 * Điền thông tin vào ô textbox
	 * Parameter: String locator: định danh, 
	 * 			  String message: thông tin cần điền
	 */
	public void sendKey(String locator, String message, int timeSleep) throws Exception{
		getElenment(locator).clear();
		//sleep(100);
		getElenment(locator).sendKeys(message);
		sleep(timeSleep);
	}
	
	/*
	 * Author: hieuht
	 * Chọn thông tin vào dropdownlist
	 * Parameter: String locator: định danh, 
	 * 			  String message: thông tin cần chọn
	 */
	public void selectByVisibleText(String locator, String message, int timeSleep) throws Exception{
		new Select(getElenment(locator)).selectByVisibleText(message);
		sleep(timeSleep);
	}
	
	/*
	 * Author: hieuht
	 * Kiểm tra định danh đã được chọn chưa
	 * Parameter: String locator: định danh
	 * Output: boolean: true/false
	 */
	public boolean isSelected(String locator){
		if(getElenment(locator).isSelected())
			return true;
		else
			return false;
	}
	
	/*
	 * Author: hieuht
	 * Lấy text từ dropdownlist
	 * Parameter: String locator: định danh
	 * Output: String: text lấy được
	 */
	public String getTextFromDropDownList(String locator){
		Select select = new Select(getElenment(locator));
		WebElement option = select.getFirstSelectedOption();
		return option.getText();
	}
	
	/*
	 * Author: hieuht
	 * Xóa dữ liệu trong textbox
	 * Parameter: String locator: định danh
	 */
	public void clear(String locator) throws Exception{
		getElenment(locator).clear();
		//sleep(500);
	}
	
	/*
	 * Author: hieuht
	 * Chuyển đổi tháng từ chuỗi sang số
	 * Parameter: String month: chuỗi (tháng)
	 * Output: int: số (tháng)
	 */
	public int convertMonthToNumber(String month){
		switch(month){
		case "January":  return 1;
		case "February":  return 2;
		case "March":  return 3;
		case "April":  return 4;
		case "May":  return 5;
		case "June":  return 6;
		case "July":  return 7;
		case "August":  return 8;
		case "September":  return 9;
		case "October": return 10;
		case "November": return 11;
		case "December": return 12;
		default: return -1;
	}
	}
	
	/*
	 * Author: hieuht
	 * Lấy text từ textbox
	 * Parameter: String locator: định danh
	 * Output: String: text lấy được
	 */
	public String getText(String locator){
		return getElenment(locator).getText();
	}

	/*
	 * Author: hieuht
	 * Chọn ngày tháng trong datetimepicker
	 * Parameter: String locator: định danh, 
	 * 			  String message: chuỗi ngày tháng
	 */
	public void pickDate(String locator, String message, int timeSleep) throws Exception{		
		Date javaDate= DateUtil.getJavaDate(Double.parseDouble(message));
		Calendar cal = Calendar.getInstance();
		cal.setTime(javaDate);
		
		this.click(locator, 100);
		int dayWant = cal.get(Calendar.DAY_OF_MONTH);
		int monthWant = cal.get(Calendar.MONTH)+1;	
		int yearWant = cal.get(Calendar.YEAR);
		
		int currentYear = Integer.parseInt(this.getText("className=ui-datepicker-year"));
		
		boolean found = false;	
		while(!found){
			if(currentYear > yearWant){
				this.click("xpath=.//*[@id='ui-datepicker-div']/div/a[1]/span", 500);
				currentYear = Integer.parseInt(this.getText("className=ui-datepicker-year"));
			}else if(currentYear < yearWant){
				this.click("xpath=.//*[@id='ui-datepicker-div']/div/a[2]/span", 500);
				currentYear = Integer.parseInt(this.getText("className=ui-datepicker-year"));
			}else if(currentYear == yearWant){
				found = true;
			}
		}
		found = false;
		String month = this.getText("className=ui-datepicker-month");
		int currentMonth = this.convertMonthToNumber(month);
		
		while(!found){	
			if(currentMonth > monthWant){
				this.click("xpath=.//*[@id='ui-datepicker-div']/div/a[1]/span", 500);
				int compareMonth = this.convertMonthToNumber(this.getText("className=ui-datepicker-month"));
				currentMonth = compareMonth;
			}else if(currentMonth < monthWant){
				this.click("xpath=.//*[@id='ui-datepicker-div']/div/a[2]/span", 500);
				int compareMonth = this.convertMonthToNumber(this.getText("className=ui-datepicker-month"));
				currentMonth = compareMonth;
			}else if(currentMonth == monthWant){
				click("linkText="+dayWant, 500);
				found = true;
			}
		}
		sleep(timeSleep);
	}
	
	/*
	 * Author: hieuht
	 * Lấy link website đang ở
	 * Output: String: chuỗi (link)
	 */
	public String getCurrentUrl() {
		return driver.getCurrentUrl();
	}
	
	/*
	 * Author: hieuht
	 * Đóng trình duyệt firefox
	 */
	public void close() throws IOException{
		this.driver.close();
		
	}
	
	/*
	 * Author: hieuht
	 * Tự động điền vào form theo action và type
	 * Parameter: ArrayList<ObjectInput> array: mảng chứa đối tượng ObjectInput
	 * 			  int column: cột lấy dữ liệu
	 * 			  ExcelAction excel: file excel chứa dữ liệu
	 */
	public void autoFill(ObjectInput[] array, int column, ExcelAction excel) throws Exception{
		for(int i=0; i<array.length; i++){
			if(array[i].getActon() == 0)
				continue;
			if(array[i].getActon() == 1)
				doAction(array[i].getType(), array, column, excel, i);
			if(array[i].getActon() == 2){
				if(excel.getStringData(column, i).equals(""))
					continue;
				else if(excel.getStringData(column, i).equals("#SPACE#")) {
					excel.printStringIntoExcel(column, i, "");
					doAction(array[i].getType(), array, column, excel, i);
				} else {
					doAction(array[i].getType(), array, column, excel, i);
				}
			}
		}
	}
	
	/*
	 * Author: hieuht
	 * Thực hiện hành động dựa vào type
	 * Parameter: int number: số action
	 * 			  ArrayList<ObjectInput> array: mảng chứa đối tượng ObjectInput
	 * 			  int column: cột lấy dữ liệu
	 * 			  ExcelAction excel: file excel chứa dữ liệu
	 */
	private void doAction(int number, ObjectInput[] array, int column, ExcelAction excel, int i) throws Exception{
		switch(number){
			case 1:{	//case 1 là selectvisible
				selectByVisibleText(array[i].getId(), excel.getStringData(column, i), array[i].getTimeSleep());
				break;
			}
			case 2:{	//case 2 là sendKey
				sendKey(array[i].getId(), excel.getStringData(column, i), array[i].getTimeSleep());
				break;
			}
			case 3:{	//case 3 là pick date
				pickDate(array[i].getId(), excel.getStringData(column, i), array[i].getTimeSleep());
				break;
			}
			case 4:{	//case 4 là click
				click(array[i].getId(), array[i].getTimeSleep());
				break;
			}
			case 5:{	//case 5 là chọn text ở dropdownlist theo chỉ số
				new Select(getElementFormList(array[i].getId(),array[i].getIndex())).selectByVisibleText(excel.getStringData(column, i));
				sleep(array[i].getTimeSleep());
				break;
			}
			case 6:{	//case 6 là click vào ô theo chỉ số
				getElementFormList(array[i].getId(), array[i].getIndex()).click();
				sleep(array[i].getTimeSleep());
				break;
			} 
			case 7:{	//click chọn hoặc không chọn
				if(excel.getStringData(column, i).equals("1")){
					if(!getElementFormList(array[i].getId(), array[i].getIndex()).isSelected()){
						getElementFormList(array[i].getId(), array[i].getIndex()).click();
						sleep(array[i].getTimeSleep());
						break;
					}
				}else if(excel.getStringData(column, i).equals("0")){
					if(getElementFormList(array[i].getId(), array[i].getIndex()).isSelected()){
						getElementFormList(array[i].getId(), array[i].getIndex()).click();
						sleep(array[i].getTimeSleep());
						break;
					}
				}
				break;
			}
			case 8:{	//click dùng 2 thuộc tính
				if(excel.getStringData(column, i).equals("1")){
					if(getElenment(array[i].getName()).getAttribute("checked") == null){
						getElementFormList(array[i].getId(), array[i].getIndex()).click();
						sleep(array[i].getTimeSleep());
						break;
					}
				}else if(excel.getStringData(column, i).equals("0")){
					if(getElenment(array[i].getName()).getAttribute("checked") != null){
						getElementFormList(array[i].getId(), array[i].getIndex()).click();
						sleep(array[i].getTimeSleep());
						break;
					}
				}
				break;
			}
		}
	}
	
	 /*Tắt alert hiển thị và lấy text của nó */
    public boolean acceptNextAlert = true;   
    public String closeAlertAndGetItsText(int timeSleep) throws Exception{
    	try {
          Alert alert = driver.switchTo().alert();
          String alertText = alert.getText();
          if (acceptNextAlert) {
            alert.accept();
          } else {
            alert.dismiss();
          }
          return alertText;
        } finally {
          acceptNextAlert = true;
          sleep(timeSleep*2);
        }
      }
	
	/*
	 * Author: hieuht
	 * Mở trình duyệt firefox và redirect tới link, set time đợi
	 * Parameter: String url: địa chỉ link, 
	 * 			  int timeSleep: time đợi
	 */
	public void open(String url) throws Exception{
		this.driver.get(url);
	}
	
	/*
	 * Author: hieuht
	 * Fill data to form login
	 * Parameter: String locatorUser: định danh ô UserName
	 * 			  String userName: tên username
	 * 			  String locatorPass: định danh ô password
	 * 			  String password: password
	 * 			  String locatorSubmit: định danh ô submit
	 */
	public void login(String link, String locatorUser, String userName, String locatorPass, String password, String locatorSubmit) throws Exception{
		this.open(link);
		this.sendKey(locatorUser, userName, 0);
	    this.sendKey(locatorPass, password, 0);
	    this.click(locatorSubmit, 0);	
	}
	
	/*
	 * Author: hieuht
	 * Chuyển hướng tới link
	 * Parameter: String url: link cần đến
	 */
	public void get(String url) throws Exception{
		driver.get(url);
	}
	
	/*
	 * Author: hieuht
	 * Nghỉ theo time (Tính thời gian nghỉ)
	 * Parameter: int timeSleep : thời gian chờ
	 */
	public void sleep(int timeSleep) throws Exception{
		countTime += timeSleep;
		Thread.sleep(timeSleep);
	}

	public int getCountTime() {
		return countTime;
	}

	public void setCountTime(int countTime) {
		this.countTime = countTime;
	}
}
