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

		// �������ݿ�����
		Connection con = null;
		Statement stmt = null;
		ResultSet rs = null;

		String url = "jdbc:sqlserver://127.0.0.1:1434;databaseName=excel;user=sa;password=3889187";// sa�������

		try {
			// Establish the connection.
			System.out.println("begin.");
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			con = DriverManager.getConnection(url);
			System.out.println("end.");
			
			
			String[][] strArr = getData();
			
			

			// Create and execute an SQL statement that returns some data.
			for (int i = 1; i < rowCount; i++) {
				String SQL = "insert into coursePlan(�꼶,רҵ,רҵ����,�γ�����,ѡ������,ѧ��,ѧʱ,ʵ��ѧʱ,�ϻ�ѧʱ,��������,�ον�ʦ,��ע) values(?,?,?,?,?,?,?,?,?,?,?,?)";

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

	// ����EXCEL������EXCEL������
	private static String[][] getData() {
		// TODO Auto-generated method stub
		String[][] data;// ��Ŵ�EXCEL
		data = new String[1000][12];

		// ��������EXCEL

		try {
			// ����workbook������
			Workbook workbook = Workbook.getWorkbook(new File("f://�����.xls"));
			// ��ȡ��һ��������sheet
			Sheet sheet = workbook.getSheet(0);
			// ��ȡ����
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
