/*
 * Lớp thực hiện các tác động tới file excel input và output
 * @Author: hieuht
 * @Date: 16/09/2016
 */
package Common;
     
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

public class ExcelAction {
	HSSFWorkbook wb = null;
	HSSFSheet sheet = null;
	FileInputStream file = null;
	String pathFile = "";
	String pathOutput = "";
	
	/*
	 * Author: hieuht
	 * Parameter: String pathInputFile: path file input
	 */
	public ExcelAction(String pathfile, String buildID, String timeBuild) throws Exception{
		pathFile = pathfile;
		pathOutput = pathFile.substring(0, pathFile.length()-4)  + "-Build#"+buildID+"-Time#"+timeBuild+"-Output.xls";
		file = new FileInputStream(new File(pathOutput));
		wb = new HSSFWorkbook(file);
	}
	
	public ExcelAction() {
		
	}
	
	public ExcelAction(String pathfile) throws Exception {
		file = new FileInputStream(new File(pathfile));
		wb = new HSSFWorkbook(file);
	}
	
	public void createFileOutput(String pathfile, String buildID, String timeBuild) throws Exception {
		pathFile = pathfile;
		pathOutput = pathFile.substring(0, pathFile.length()-4) + "-Build#"+buildID+"-Time#"+timeBuild+"-Output.xls";
		file = new FileInputStream(new File(pathfile));
		wb = new HSSFWorkbook(file);
		
		FileOutputStream outFile =new FileOutputStream(new File(pathOutput));
		wb.write(outFile);
		outFile.close();
		
		file = new FileInputStream(new File(pathOutput));
		wb = new HSSFWorkbook(file);
	}
	
	/*
	 * Author: hieuht
	 * Truy cập vào file excel input
	 * Parameter: int NumberSheet: số sheet cần truy cập trong file input
	 */
	public void accessSheet(String sheetName) throws Exception{
		sheet = wb.getSheet(sheetName);
	}
	
	public HSSFSheet getSheet() throws Exception{
		return this.sheet;
	}
	
	/*
	 * Author: hieuht
	 * Ghi file output, đóng file output, đóng file input
	 */
	public void finish() throws Exception{
		file.close();
		write();
		wb.close();
	}
	
	public void write() throws Exception{
		FileOutputStream outFile =new FileOutputStream(new File(pathOutput));
		wb.write(outFile);
		outFile.close();
	}
	
	public short getColumn(String nameColumn){
		CellReference cell = new CellReference(nameColumn);
		return cell.getCol();
	}
	
	/*
	 * Author: hieuht
	 * Lấy dữ liệu chuỗi dựa vào dòng, cột, file (input/output), mục đích sử dụng (input/output)
	 * Parameter: int column: số cột, 
	 * 			  int row: số dòng, 
	 * Output: String: chuỗi nhận được
	 */
	public String getStringData(int column, int row) throws Exception{
		try{
			Cell cell = sheet.getRow(row).getCell(column);
			
			if(cell == null)
				return "";
			else {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				return Common.checkNull(cell.getStringCellValue());	
			}
		}catch(NullPointerException e){
			return "";
		}
	}
	
	public String getFormulaCellData(int column1, int row1) throws Exception{
		CellReference cellReference = new CellReference(row1, column1); 
		Row row = sheet.getRow(cellReference.getRow());
		Cell cell = row.getCell(cellReference.getCol());
		FormulaEvaluator formulaEval = wb.getCreationHelper().createFormulaEvaluator();
		String value=formulaEval.evaluate(cell).formatAsString();
		return value.substring(1, value.length()-1);
	}
	
	/*
	 * Author: hieuht
	 * In kết quả vào excel output
	 * Parameter: int column: số hàng, 
	 * 			  int row: số cột, 
	 * 			  boolean result: True/False
	 */
	public void printResultIntoExcel(int column, int row, boolean result) throws Exception{
		HSSFRow rows     = sheet.getRow((short)row); 
		HSSFCell cells   = rows.createCell((short)column); 
		if(result == true)
			cells.setCellValue("SUCCESSED"); 
		else
			cells.setCellValue("FAILED"); 
	}
	
	public void printStringIntoExcel(int column, int row, String string) throws Exception{
		HSSFRow rows     = sheet.getRow((short)row); 
		HSSFCell cells   = rows.createCell((short)column); 
		cells.setCellValue(string); 
	}
	
	
	
	/*
	 * Author: hieuht
	 * So sánh 2 chuỗi và in kết quả vào excel output
	 * Parameter: String string1: chuỗi 1, 
	 * 			  String string2: chuỗi 2,
	 * 			  int column: số hàng, 
	 * 			  int row: số cột
	 */
	public void compareStringAndPrint(String string1, String string2, int column, int row) throws Exception{
		if(string1.equals(string2))
			this.printResultIntoExcel(column, row, true);
		else
			this.printResultIntoExcel(column, row, false);
	}
	
	public void printResultToFlow(String result, String compareString, ExcelAction excel, int row, int column) throws Exception{
		if(result.equals(compareString))
			printResultIntoExcel(column, row, true);
		else
			printResultIntoExcel(column, row, false);
	}
}
