package DonPC;

import java.sql.ResultSet;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;

import Common.DatabaseAction;
import Common.ExcelAction;
import Common.ObjectInput;

public class CheckDatabase_PC {
	private Logger logger;
	public CheckDatabase_PC(){
		//Khởi tạo logger
		logger=Logger.getLogger("check Database");
		//Config log4j
		PropertyConfigurator.configure("log4j.properties");
	}
	/*
	 * Author: hieuht
	 * So sánh dữ liệu trong db và excel bảng proposals
	 * Input: ExcelAction excel: file excel, 
	 * 		  int column: cột dữ liệu, 
	 * 		  int row: hàng dữ liệu, 
	 * 		  ResultSet rs: dữ liệu truy vấn, 
	 * Output: boolean:  true/false 
	 */
	public boolean checkDataInTableProposals(ExcelAction excel, int column, int row, ResultSet rs) throws Exception {
		//1 = int, 2 = string, 3 = cut datetime
		ObjectInput[] array = {	new ObjectInput("id", 2),
								new ObjectInput("creator_id",2),
								new ObjectInput("type_id", 2),
								new ObjectInput("helped_by",2)
		};
		return DatabaseAction.compare(array, logger, excel, column, row, rs);
	}
	
	/*
	 * Author: hieuht
	 * So sánh dữ liệu trong db và excel bảng pc_proposal
	 * Input: ExcelAction excel: file excel, 
	 * 		  int column: cột dữ liệu, 
	 * 		  int row: hàng dữ liệu, 
	 * 		  ResultSet rs: dữ liệu truy vấn, 
	 * Output: true/false 
	 */
	public boolean checkDataInTablePcProposal(ExcelAction excel, int column, int row, ResultSet rs) throws Exception {
		//1 = int, 2 = string, 3 = cut datetime	
		ObjectInput[] array = {	new ObjectInput("proposal_id", 2),
								new ObjectInput("proposal_code", 2),			
								new ObjectInput("creator_id", 2),
								new ObjectInput("helped_by", 2),	
								new ObjectInput("proposal_date", 3),
								new ObjectInput("company_id", 2),
								new ObjectInput("group_id", 2),
								new ObjectInput("division_id", 2),
								new ObjectInput("pc_user_use", 2),
								new ObjectInput("pc_category", 2),
								new ObjectInput("pc_code", 2),	
								new ObjectInput("target_proposal", 2),			
								new ObjectInput("status_permission_admin", 2),
								new ObjectInput("status_access_network", 2),
								new ObjectInput("status_limit_usb", 2),
								new ObjectInput("status_outside_laptop", 2),
								new ObjectInput("status_limit_sd", 2),
								new ObjectInput("status_other", 2),
								new ObjectInput("note_other", 2),
								new ObjectInput("start_maturity_date", 3),
								new ObjectInput("end_maturity_date", 3),	
								new ObjectInput("work_place", 2),			
								new ObjectInput("appointed_date", 3),
								new ObjectInput("comment", 2),
								new ObjectInput("reason_maturity", 2),
								new ObjectInput("status", 2),
								new ObjectInput("re_proposal", 2),
								new ObjectInput("last_operator", 2),
								new ObjectInput("user_id_back", 2),
								new ObjectInput("user_id_change_process", 2),
								new ObjectInput("user_id_change_complete", 2),
								new ObjectInput("execute_date", 3),
								new ObjectInput("create_for_manager", 2),
		};
		return DatabaseAction.compare(array, logger, excel, column, row, rs);
	}
	
	/*
	 * Author: hieuht
	 * So sánh dữ liệu trong db và excel bảng ts_proposal_histories
	 * Input: ExcelAction excel: file excel, 
	 * 		  int column: cột dữ liệu, 
	 * 		  int row: hàng dữ liệu, 
	 * 		  ResultSet rs: dữ liệu truy vấn, 
	 * Output: true/false 
	 */
	public boolean checkDataInTablePcProposalHistories(ExcelAction excel, int column, int row, ResultSet rs) throws Exception {
		//1 = int, 2 = string, 3 = cut datetime
		ObjectInput[] array = {	new ObjectInput("proposal_id", 2),
								new ObjectInput("user_id", 2),
								new ObjectInput("action", 2),
								new ObjectInput("groupset", 2),
								new ObjectInput("comment", 2),
								new ObjectInput("action_content", 2)
		};
		return DatabaseAction.compare(array, logger, excel, column, row, rs);
	}
	
	/*
	 * Author: hieuht
	 * So sánh dữ liệu trong db và excel bảng pc_proposal_process
	 * Input: ExcelAction excel: file excel, 
	 * 		  int column: cột dữ liệu, 
	 * 		  int row: hàng dữ liệu, 
	 * 		  ResultSet rs: dữ liệu truy vấn, 
	 * Output: true/false 
	 */
	public boolean checkDataInTablePcProposalProcess(ExcelAction excel, int column, int row, ResultSet rs) throws Exception {
		//1 = int, 2 = string, 3 = cut datetime
		ObjectInput[] array = {	new ObjectInput("creator_id", 2),
								new ObjectInput("proposal_id", 2),
								new ObjectInput("user_inspect_id", 2),
								new ObjectInput("user_infor_id", 2),
								new ObjectInput("approval_status", 2),
								new ObjectInput("status_all", 2),
								new ObjectInput("groupset", 2),
								new ObjectInput("group_title", 2),
								new ObjectInput("comment", 2)
		};
		return DatabaseAction.compare(array, logger, excel, column, row, rs);
	}
}
