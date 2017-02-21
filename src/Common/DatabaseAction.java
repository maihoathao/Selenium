/*
 * Lớp thực hiện các tác động chung của chương trình tới db
 * @Author: hieuht
 * @Date: 16/09/2016
 */
package Common;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;

import org.apache.log4j.Logger;

import com.jcraft.jsch.JSch;
import com.jcraft.jsch.JSchException;
import com.jcraft.jsch.Session;
import com.mysql.jdbc.jdbc2.optional.MysqlDataSource;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;

public class DatabaseAction {
	/*
	 * Statement:
	 * Connection:
	 * dbuserName: username kết nối db
	 * dbpassword: password kết nối đb
	 * database: tên đb kết nối
	 * inputFile: path file excel đọc dữ liệu
	 * nLocalPort: port
	 * pathMySQL: đường dẫn mysql
	 */
	public static Statement stmt = null;
	public static Connection connection = null;
	private static String dbuserName = "";
	private static String dbpassword = "";
	private static String database = "";
	private static String inputFile;
	private static String nLocalPort = "";
	private static String pathMySQL = "";
	/*
	 * Author: hieuht
	 * Setup pahFileInput để đọc lấy dữ liệu trong excel
	 * Parameter: String inputFile: đường dẫn file input
	 */
	public static void setInputFile(String inputFile) {
		DatabaseAction.inputFile = inputFile;
	}

	/*
	 * Author: hieuht
	 * Chọn môi trường để chương trình kết nối database
	 * Parameter: String Env: tên môi trường (Dev/Test/Staging/Live)
	 */
	
	public static void doSshTunnel(String Env, int count, String buildId, String timeBuild) throws Exception{
		//Lấy sheet 0 từ file excel input
		ExcelAction excel = new ExcelAction(inputFile, buildId, timeBuild);
		excel.accessSheet("Config");
		
		//Case môi trường = Dev
		if(Env.equals("Dev")){
			//Đọc cái config trong sheet 0
    		dbuserName = excel.getStringData(6, 2);
		    dbpassword = excel.getStringData(7, 2);
		    database = excel.getStringData(8, 2);
		    nLocalPort = excel.getStringData(5, 2);
		    pathMySQL = excel.getStringData(14, 2);
		    //Kết nối DB
    		Class.forName("com.mysql.jdbc.Driver").newInstance();
    		
    		connection = DriverManager.getConnection(pathMySQL+nLocalPort+"/"+database+"?useUnicode=true&characterEncoding=UTF-8&charSet=UTF-8&zeroDateTimeBehavior=convertToNull", dbuserName, dbpassword);
    		stmt = connection.createStatement(); 
	    }else{
	    	String privateKey = "keySsh\\aitdev";
	        String strSshUser = "aitdev"; // SSH loging username
	        String strSshPassword = "rEWGtRZ0n?Q?u,\\vo#r7"; // SSH login password
	        String strSshHost = "27.110.48.29";  // hostname or ip or SSH server
	        int nSshPort = 22; // remote SSH host port number
	        String strRemoteHost = "127.0.0.1"; // hostname or ip of your database server       
	        int nRemotePort = 3306; // remote port number of your database
	        int localPort = 8740; // any free port can be used
	        String localSSHUrl = "127.0.0.1";
	        
	        final JSch jsch = new JSch();  
	        jsch.addIdentity(privateKey,"rEWGtRZ0n?Q?u,\\vo#r7");
	        Session session = jsch.getSession( strSshUser, strSshHost, 22 );  
	        session.setPassword( strSshPassword );    

	        final Properties config = new Properties();
	        config.put( "StrictHostKeyChecking", "no" );
	        config.put("ConnectionAttempts", "3");
	        session.setConfig( config );
	        
	        session.connect();
	       
	        //System.out.println("session.isConnected " + session.isConnected());
	        //session.disconnect();
	        if(count == 0)
	        	session.setPortForwardingL(8740, "10.0.33.10", 3306);	    
	          
	    	if(Env.equals("Test")){ // Connect database to environment Test
	    		dbuserName = "rgm_test";
	            dbpassword = "Apr8Httjkmo";
	            database = "ringi_test";
			    
		    }else if(Env.equals("Staging")){ // Connect database to environment Staging
		    	dbuserName = "rg2m_test";
			    dbpassword = "UKHqR9xtAHPiDidBda1F"; //Not filled because of security reasons
			    database = "ringi2_test";
			    
		    }else if(Env.equals("Live")){ // Connect database to environment Live
		    	dbuserName = "";
			    dbpassword = ""; //Not filled because of security reasons
			    database = "";
		    }
	    	
	        //mysql database connectivity
	        MysqlDataSource dataSource = new MysqlDataSource();
	        dataSource.setServerName(localSSHUrl);
	        dataSource.setPortNumber(localPort);
	        dataSource.setUser(dbuserName);
	        dataSource.setAllowMultiQueries(true);
	        //dataSource.setURL("jdbc:mysql://localhost:3306/yourdatabase?zeroDateTimeBehavior=convertToNull");
	        dataSource.setPassword(dbpassword);
	        dataSource.setDatabaseName(database);	
	        if(connection != null && stmt != null){
	        	connection.close();
	        	stmt.close();
	        }
	        connection = dataSource.getConnection();
	        stmt = connection.createStatement();
	    }			
	}
	
	/*
	 * Author: hieuht
	 * Thực hiện reset data: xóa db cũ, import db mới
	 * Parameter: String dbNameDelete: tên database muốn xóa
	 * 			  String dbNameImport: tên database muốn import
	 */
	public static void resetData(String dbNameDelete, String dbNameImport) throws Exception{
//		//Xóa database
//		deleteDabase(dbNameDelete);
//		//Import database
//		createDabase(dbNameImport);
//		
//		//Truy cập vào sheet 0 file input
//		ExcelAction excel = new ExcelAction(inputFile);
//		excel.accessSheet("Config");
//			
//		//Đọc các thông tin config
//		String pathSql = excel.getStringData(1, 14);
//		String pathCmdMysql = excel.getStringData(1, 15);
//		String user = excel.getStringData(1, 16);
//		String pass = excel.getStringData(1, 17);
//		
//		//Thực hiện rs trên comandline
//		performOnCommandLine(pathSql, pathCmdMysql, user, pass);
//		
//		Class.forName("com.mysql.jdbc.Driver").newInstance();
//		DatabaseAction.connection = DriverManager.getConnection(pathMySQL+nLocalPort+"/"+database+"?useUnicode=true&characterEncoding=UTF-8&charSet=UTF-8&zeroDateTimeBehavior=convertToNull", dbuserName, dbpassword);
//		DatabaseAction.stmt = DatabaseAction.connection.createStatement(); 
	}
	
	/*
	 * Author: hieuht
	 * Thực thi lệnh trên cmd line
	 * Parameter: String FileName: tên database import
	 * 			  String m_MySqlPath: path mysql
	 * 			  String user: username kết nối mysql
	 * 			  String pass: pass kết nối mysql
	 */
	private static void performOnCommandLine(String FileName, String m_MySqlPath, String user, String pass) throws InterruptedException, IOException{		
		String[] command = new String[]{"\""+m_MySqlPath+"mysql\"", database, "-u" + user, "-e", " source "+"\""+FileName+"\"" };
		Process runtimeProcess = Runtime.getRuntime().exec(command, null, new File(m_MySqlPath));
		runtimeProcess.waitFor();
	}
	/*
	 * Author: hieuht
	 * So sánh dữ liệu trong db và mảng
	 * Parameter: ObjectInput[] array: mảng chứa dữ liệu với tên và action (chia các case theo action)
	 * 			  Logger logger: ghi ra file log 
	 *			  ExcelAction excel: file excel chứa dữ liệu 
	 *			  int column: chỉ số cột lấy dữ liệu 
	 *			  int row: chỉ số cột lấy dữ liệu
	 *			  ResultSet rs: kết quả truy vấn db
	 * Output: boolean: true/false
	 */
	public static boolean compare(ObjectInput[] array, Logger logger, ExcelAction excel, int column, int row, ResultSet rs) throws Exception{
		for(int z = 0; z< array.length; z++){
			if(array[z].getType() == 2){
				logger.debug("Check column name:  " + array[z].getName());
				if(!DatabaseAction.compareString(excel.getStringData(column, row), DatabaseAction.getStringData(rs, array[z].getName()))){
					logger.error("Fail in column name:  " + array[z].getName());
					return false;
				}
				row+=1;
			}else if(array[z].getType() == 3){		
				logger.debug("Check column name:  " + array[z].getName());
				if(DatabaseAction.getStringData(rs, array[z].getName()).equals("")){
					if(!DatabaseAction.compareString(excel.getStringData(column, row), DatabaseAction.getStringData(rs, array[z].getName()))){
						logger.error("Fail in column name:  " + array[z].getName());
						return false;
					}
				}else{
					if(!DatabaseAction.compareString(excel.getStringData(column, row), DatabaseAction.getStringData(rs, array[z].getName()).substring(0, 10))){
						logger.error("Fail in column name:  " + array[z].getName());
						return false;
					}
				}
				row+=1;
			}
		}
		return true;
	}
	
	/*
	 * Author: hieuht
	 * Xóa db theo tên
	 * Parameter: String dbName: tên db muốn xóa
	 */
	public static void deleteDabase(String dbName){
		try {
			stmt.executeUpdate("DROP DATABASE "+dbName);
		} catch (SQLException e) {
			System.exit(0);
		}
	}
	
	/*
	 * Author: hieuht
	 * Tạo db theo tên
	 * Parameter: String dbName: tên db muốn tạo
	 */
	public static void createDabase(String dbName){
		try {
			stmt.executeUpdate("CREATE DATABASE "+dbName);
		} catch (SQLException e) {
			System.exit(0);
		}
	}

	/*
	 * Author: hieuht
	 * Lấy dữ liệu kiểu int từ kết quả truy vấn
	 * Parameter: ResultSet rs: kết quả truy vấn db
	 * 			  String nameColumn: tên trường cần lấy dữ liệu
	 * Output: String: dữ liệu lấy được
	 */
	public static String getStringData(ResultSet rs, String nameColumn) throws NumberFormatException, SQLException{
		return Common.checkNull(rs.getString(nameColumn));
	}
	
	public static void closeConnect() throws Exception{
		stmt.close();
		connection.close();
	}
		
	/*
	 * Author: hieuht
	 * So sánh 2 chuỗi truyền vào
	 * Parameter: String1: chuỗi 1
	 * 			  String2: chuỗi 2
	 * Output: boolean: true/flase
	 */
	public static boolean compareString(String string1, String string2){
		return string1.equals(string2) ?  true :  false;
	}
}
