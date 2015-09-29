package SQL;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class sqlHelper {

	static int rowCount = 0;

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		// 定义数据库所需
		Connection con = null;
		Statement stmt = null;
		ResultSet rs = null;

		String url = "jdbc:sqlserver://127.0.0.1:1434;databaseName=excel;user=sa;password=3889187";// sa身份连接

		try {
			// Establish the connection.
			System.out.println("begin.");
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			con = DriverManager.getConnection(url);
			System.out.println("end.");
			
			
			String[][] strArr = getData();
			
			

			// Create and execute an SQL statement that returns some data.
			for (int i = 1; i < rowCount; i++) {
				String SQL = "insert into coursePlan(年级,专业,专业人数,课程名称,选修类型,学分,学时,实验学时,上机学时,起迄周序,任课教师,备注) values(?,?,?,?,?,?,?,?,?,?,?,?)";

				PreparedStatement ps = con.prepareStatement(SQL);
				for(int j=0;j<12;++j){
					
					ps.setString(j+1,strArr[i][j]);
				}
				boolean flag = ps.execute();

			}
		}

		// Handle any errors that may have occurred.
		catch (Exception e) {
			e.printStackTrace();
		}

		finally {
			if (rs != null)
				try {
					rs.close();
				} catch (Exception e) {
				}
			if (stmt != null)
				try {
					stmt.close();
				} catch (Exception e) {
				}
			if (con != null)
				try {
					con.close();
				} catch (Exception e) {
				}
		}
	}

	// 解析EXCEL并返回EXCEL的数据
	private static String[][] getData() {
		// TODO Auto-generated method stub
		String[][] data;// 存放从EXCEL
		data = new String[1000][12];

		// 解析读入EXCEL

		try {
			// 创建workbook工作薄
			Workbook workbook = Workbook.getWorkbook(new File("f://计算机.xls"));
			// 获取第一个工作表sheet
			Sheet sheet = workbook.getSheet(0);
			// 获取数据
			rowCount = 0;
			for (int i = 0; i < sheet.getRows(); i++) {
				for (int j = 0; j < sheet.getColumns(); j++) {
					Cell cell = sheet.getCell(j, i);
					data[i][j] = cell.getContents().trim();
				}
				rowCount++;
			}
			workbook.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return data;
	}
}
