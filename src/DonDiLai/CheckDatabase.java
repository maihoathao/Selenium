/*
 * Thực hiện các hàm check dữ liệu trong database và excel
 * @Author: hieuht
 * @Date: 16/09/2016
 */
package DonDiLai;

import Common.*;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;

public class CheckDatabase {
	private Logger logger;
	public CheckDatabase(){
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
	 * So sánh dữ liệu trong db và excel bảng ts_proposal
	 * Input: ExcelAction excel: file excel, 
	 * 		  int column: cột dữ liệu, 
	 * 		  int row: hàng dữ liệu, 
	 * 		  ResultSet rs: dữ liệu truy vấn, 
	 * Output: true/false 
	 */
	public boolean checkDataInTableTsProposal(ExcelAction excel, int column, int row, ResultSet rs) throws Exception {
		//1 = int, 2 = string, 3 = cut datetime	
		ObjectInput[] array = {	new ObjectInput("proposal_id", 2),
								new ObjectInput("creator_id", 2),			
								new ObjectInput("helped_by", 2),
								new ObjectInput("proposal_date", 3),	
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
								new ObjectInput("old_post_code", 2),
								new ObjectInput("old_address", 2),
								new ObjectInput("comment", 2),
								new ObjectInput("status", 2),
								new ObjectInput("re_proposal", 2),
								new ObjectInput("total_cost_one_day", 2),
								new ObjectInput("total_cost_one_month", 2),	
								new ObjectInput("total_cost_three_month", 2),			
								new ObjectInput("total_cost_six_month", 2),
								new ObjectInput("start_point", 2),
								new ObjectInput("end_point", 2),
								new ObjectInput("determine_content", 2),
								new ObjectInput("last_operator", 2),
								new ObjectInput("user_id_determine", 2),
								new ObjectInput("user_id_edit", 2),
								new ObjectInput("proposal_edit", 2),
								new ObjectInput("current_proposal", 2),
								new ObjectInput("parrent_proposal", 2),
								new ObjectInput("old_proposal_id", 2),
								new ObjectInput("version", 2),	
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
	public boolean checkDataInTableTsProposalHistories(ExcelAction excel, int column, int row, ResultSet rs) throws Exception {
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
	 * So sánh dữ liệu trong db và excel bảng ts_proposal_route
	 * Input: ExcelAction excel: file excel, 
	 * 		  int column: cột dữ liệu, 
	 * 		  int row: hàng dữ liệu, 
	 * 		  ResultSet rs: dữ liệu truy vấn, 
	 * Output: true/false 
	 */
	public boolean checkDataInTableTsProposalRoute(ExcelAction excel, int column, int row, ResultSet rs) throws Exception {
		//1 = int, 2 = string, 3 = cut datetime
		ObjectInput[] array = {	new ObjectInput("proposal_id", 2),
								new ObjectInput("route_names", 2),
								new ObjectInput("station_from", 2),
								new ObjectInput("station_to", 2),
								new ObjectInput("cost_of_route", 2)
		};
		return DatabaseAction.compare(array, logger, excel, column, row, rs);
	}
	
	/*
	 * Author: hieuht
	 * So sánh dữ liệu trong db và excel bảng ts_proposal_process
	 * Input: ExcelAction excel: file excel, 
	 * 		  int column: cột dữ liệu, 
	 * 		  int row: hàng dữ liệu, 
	 * 		  ResultSet rs: dữ liệu truy vấn, 
	 * Output: true/false 
	 */
	public boolean checkDataInTableTsProposalProcess(ExcelAction excel, int column, int row, ResultSet rs) throws Exception {
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
