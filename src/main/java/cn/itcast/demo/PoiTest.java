package cn.itcast.demo;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class PoiTest {

	public static void main(String[] args) throws IOException, InstantiationException, IllegalAccessException, ClassNotFoundException, SQLException {
		
		/*jdbc.driver=com.mysql.jdbc.Driver
				jdbc.url=jdbc:mysql://localhost:3306/pinyougoudb?characterEncoding=utf-8
				jdbc.username=root
				jdbc.password=root*/
		//使用jdbc链接数据库
		Class.forName("com.mysql.jdbc.Driver").newInstance();
		String url = "jdbc:mysql://localhost:3306/pinyougoudb?characterEncoding=utf-8";
		String user = "root";
		String password = "0826";
		
		Connection conn = DriverManager.getConnection(url, user,password);   
		Statement stmt = conn.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_UPDATABLE);   
		
		String sql = "SELECT * FROM tb_goods";   	//100万测试数据
		//sql = "select name,age,des from customer";   	//100万测试数据
		ResultSet rs = stmt.executeQuery(sql);  						//bug 要分次读取，否则记录过多
		
		Workbook book = new HSSFWorkbook(); //65536超过该数只用XSSF
		Sheet sheet = book.createSheet();
		int rowIndex = 0;
		//先创建title
		Row titleRow = sheet.createRow(rowIndex);
		Cell titleCell1 = titleRow.createCell(0);
		titleCell1.setCellValue("id");
		
		//创建单元格样式
		CellStyle style = book.createCellStyle();	
		Font font = book.createFont();
		font.setBold(true);
		font.setFontHeightInPoints((short)22); //设置字体大小
		font.setFontName("黑体");
		// 横向居中
		style.setAlignment(CellStyle.ALIGN_CENTER); 
		// 纵向居中
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER); 
		style.setFont(font);
		
		titleCell1.setCellStyle(style);
		
		
		Cell createCel2 = titleRow.createCell(1);
		createCel2.setCellStyle(style);
		createCel2.setCellValue("商家id");
		
		Cell createCel3 = titleRow.createCell(2);
		createCel3.setCellStyle(style);
		createCel3.setCellValue("商品名称");
		
		
		rowIndex++;
		
		//循环数据
		
		
		while(rs.next()) {
			Row row = sheet.createRow(rowIndex++);
			//row是根据返回的list的size进行循环
			Cell cell = row.createCell(0);
			//根据集合中对象的属性进行循环
			cell.setCellValue(rs.getString(1));
			
			Cell cell1 = row.createCell(1);
			cell1.setCellValue(rs.getString(2));
			
			Cell cell2 = row.createCell(2);
			cell2.setCellValue(rs.getString(3));
		}
		
		FileOutputStream stream = new FileOutputStream(new File("d:/test.xls"));
		book.write(stream);
		stream.close();
		System.out.println("执行完毕");
	}
}
