/*
 * Thực hiện các hàm chung của tất cả các class
 * @Author: hieuht
 * @Date: 10/10/2016
 */
package Common;

import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.HashSet;
import java.util.Hashtable;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import org.openqa.selenium.WebElement;

public class Common {
	public static Hashtable listException = new Hashtable();
	private static String fileException = "DataTest\\Define_Exception.xls";
	
	/*
	 * Author: hieuht
	 * Đọc file excel định nghĩa exception
	 * Date: 03/02/2017
	 */
	public static void readExcelDefine () throws Exception {
		ExcelAction excelReadException = new ExcelAction(fileException);
		excelReadException.accessSheet("Define");
		for(int i=1; i<excelReadException.getSheet().getPhysicalNumberOfRows(); i++) {
			listException.put(excelReadException.getStringData(1, i), excelReadException.getStringData(0, i));
		}
	}
	
	/*
	 * Author: hieuht
	 * Lấy chỉ số trong mảng dựa vào tên
	 * Parameter: ArrayList<ObjectInput> array: mảng chứa đối tượng ObjectInput
	 * 			  String name: tên
	 */
	public static int getIndexFromArray(ObjectInput[] array, String name){
		for(int j=0; j<array.length; j++){
			if(array[j].getName().equals(name))
				return j;
		}
		return -1;
	}
		
	/*
	 * Author: hieuht
	 * Kiểm tra rỗng, nếu dùng để nhập thì không so sánh, dùng để check thì trả về "111111111" nếu rỗng
	 * Parameter: String string: chuỗi kiểm tra
	 * Output: String: chuỗi sau check
	 */
	public static String checkNull(String string){
		if(string == null || string.equals("") ){
			return "";
		}
		return string;
	}
	
	/*
	 * Author: hieuht
	 * set Id ObjectInput = userName bảng users (dùng cho setting duyệt)
	 * Parameter: 
		 ObjectInput[] array: mảng chứa object 
		 String nameInArray: tên object set id
		 String userName: username truy vấn
		 String idWeb: set id
	 * 
	 */
	public static void setIdByUserName(ObjectInput[] array, String nameInArray, String userName, String idWeb)throws Exception{
		ResultSet rs = null;
		rs = DatabaseAction.stmt.executeQuery("SELECT * FROM `users` WHERE username = "+userName);
		rs.next();
		String idQuery = rs.getString("id");
		array[Common.getIndexFromArray(array, nameInArray)].setId(idWeb+idQuery);
	}
}
