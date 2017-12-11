package net.intelink.excelpoi.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.Driver;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Iterator;

import org.apache.log4j.Logger;
import org.apache.log4j.helpers.Loader;
import org.apache.log4j.rewrite.RewriteAppender;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.impl.PublicImpl;

public class PoiTest {
	static Logger logger = Logger.getLogger(PoiTest.class);

	public static void main(String[] args) throws IOException, ClassNotFoundException, SQLException {
		
	}
	
	public void test1() throws IOException {
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
		HSSFSheet sheet = hssfWorkbook.createSheet("sheet0");
		HSSFRow row = sheet.createRow(0);
		HSSFCell cell = row.createCell(0);
		cell.setCellValue("测试Poi");
		FileOutputStream fileOutputStream = new FileOutputStream("C:/Users/intelink012/Desktop/poi.xls");
		hssfWorkbook.write(fileOutputStream);
		fileOutputStream.flush();
		fileOutputStream.close();
	}

	public void test2() throws IOException {
		// 创建excel的文档对象
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
		// 创建excel的表单（sheet）
		HSSFSheet sheet = hssfWorkbook.createSheet("成绩表");
		// 在表单（sheet）里创建第一行，参数为行的索引
		HSSFRow row = sheet.createRow(0);
		// 创建单元格
		HSSFCell cell = row.createCell(0);
		// 设置单元格的内容
		cell.setCellValue("学员成绩一览表");
		// 合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));

		// 创建第二行
		HSSFRow row2 = sheet.createRow(1);
		// 创建单元格并设置内容
		row2.createCell(0).setCellValue("姓名");
		row2.createCell(1).setCellValue("班级");
		row2.createCell(2).setCellValue("笔试成绩");
		row2.createCell(2).setCellValue("机试成绩");

		HSSFRow row3 = sheet.createRow(2);
		row3.createCell(0).setCellValue("李明");
		row3.createCell(1).setCellValue("As178");
		row3.createCell(2).setCellValue("87");
		row3.createCell(3).setCellValue("78");

		HSSFRow row4 = sheet.createRow(3);
		row4.createCell(0).setCellValue("张飞");
		row4.createCell(1).setCellValue("As255");
		row4.createCell(2).setCellValue("78");
		row4.createCell(3).setCellValue("90");

		HSSFRow row5 = sheet.createRow(4);
		row5.createCell(0).setCellValue("王菲");
		row5.createCell(1).setCellValue("As336");
		row5.createCell(2).setCellValue("82");
		row5.createCell(3).setCellValue("69");

		// 使用文件输出流输出文件
		FileOutputStream fileOutputStream = new FileOutputStream(new File("C:/Users/intelink012/Desktop/poi.xls"));
		hssfWorkbook.write(fileOutputStream);
	}

	public void test3() throws IOException {
		FileInputStream fileInputStream = new FileInputStream(new File("C:/Users/intelink012/Desktop/poi.xls"));
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fileInputStream);
		HSSFSheet sheetAt = hssfWorkbook.getSheetAt(0);
		for (Row r : sheetAt) {
			if (r.getRowNum() < 2) {
				continue;
			}
			ScoreInfo scoreInfo = new ScoreInfo();
			scoreInfo.setName(r.getCell(0).getStringCellValue());
			scoreInfo.setBanji(r.getCell(1).getStringCellValue());
			scoreInfo.setCgrade(Integer.valueOf(r.getCell(2).getStringCellValue()));
			scoreInfo.setWgrade(Integer.valueOf(r.getCell(3).getStringCellValue()));
			System.out.println(r.getCell(2).getNumericCellValue());

			System.out.println(scoreInfo);
		}
	}

	public void test4() throws IOException {

		FileInputStream fis = new FileInputStream(new File("C:/Users/intelink012/Desktop/poi.xlsx"));
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheetAt = workbook.getSheetAt(0);
		Iterator<Row> iterator = sheetAt.iterator();
		while (iterator.hasNext()) {
			XSSFRow row = (XSSFRow) iterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.println(cell.getNumericCellValue() + "\t\t");
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.println(cell.getStringCellValue() + "\t\t");
				default:
					break;
				}
			}
			System.out.println();
		}
		fis.close();
	}

	public void test5() throws ClassNotFoundException, SQLException, IOException {
		// 1.链接数据库
		Class.forName("com.mysql.jdbc.Driver");
		Connection connection = DriverManager.getConnection("jdbc:mysql://localhost/oms_test?serverTimezone=UTC&useSSL=true", "root", "root");
		// 2.查询获得结果集
		Statement statement = connection.createStatement();
		ResultSet resultSet = statement.executeQuery("select * from myuser");
		// 3.创建excel表格对象
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("myuser");
		XSSFRow row = sheet.createRow(1);
		// 4.将查询到的结果写入表格中
		XSSFCell cell;
		cell=row.createCell(1);
		cell.setCellValue("Id");
		cell=row.createCell(2);
		cell.setCellValue("姓名");
		cell=row.createCell(3);
		cell.setCellValue("年龄");
		cell=row.createCell(4);
		cell.setCellValue("地址");
		cell=row.createCell(5);
		cell.setCellValue("性别");
		cell=row.createCell(6);
		cell.setCellValue("爱好");
		cell=row.createCell(7);
		cell.setCellValue("创建时间");
		
		int i=2;
		while(resultSet.next()) {
			row=sheet.createRow(i);
			
			cell=row.createCell(1);
			cell.setCellValue(resultSet.getInt(1));
			
			cell=row.createCell(2);
			cell.setCellValue(resultSet.getString(2));
			
			cell=row.createCell(3);
			cell.setCellValue(resultSet.getInt(3));
			
			cell=row.createCell(4);
			cell.setCellValue(resultSet.getString(4));
			
			cell=row.createCell(5);
			cell.setCellValue(resultSet.getString(5));
			
			cell=row.createCell(6);
			cell.setCellValue(resultSet.getString(6));
			
			cell=row.createCell(7);
			System.out.println(resultSet.getDate(7));
			cell.setCellValue(resultSet.getDate(7).toString());
			i++;
		}
		//5.将结果保存至文件中
		FileOutputStream fileOutputStream = new FileOutputStream(new File("C:/Users/intelink012/Desktop/myuser.xlsx"));
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		System.out.println("success");
	}
	
	public void test6() throws SQLException, ClassNotFoundException, FileNotFoundException, IOException {
		//1.连接数据库
		Class.forName("com.mysql.jdbc.Driver");
		Connection connection = DriverManager.getConnection("jdbc:mysql://localhost/oms_test?serverTimezone=UTC&useSSL=true", "root", "root");
		PreparedStatement statement=null;
		//2.创建excel文档对象
		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File("C:/Users/intelink012/Desktop/myuser.xlsx")));
		XSSFSheet sheetAt = workbook.getSheetAt(0);
		//3.定义sql语句
		String sql="insert into myuser(name,age,address,sex,live,creation) values(?,?,?,?,?,NOW()) ";
		//4.遍历获取一行中每个单元格的数据,将数据传入失sql语句中
		int i=2;
		int rowNum=sheetAt.getLastRowNum();

		while(i<rowNum) {
			XSSFRow row = sheetAt.getRow(i);
			String name=row.getCell(2).getStringCellValue();
			
			//int age=Integer.valueOf(row.getCell(3).getStringCellValue());
			int age=12;
		System.out.println(row.getCell(3).getNumericCellValue());
			String address=row.getCell(4).getStringCellValue();
			String sex=row.getCell(5).getStringCellValue();
			String live=row.getCell(6).getStringCellValue();
			
			statement = connection.prepareStatement(sql);
			statement.setString(1, name);
			statement.setInt(2, age);
			statement.setString(3, address);
			statement.setString(4, sex);
			statement.setString(5, live);
			//5.执行sql语句
			//statement.executeUpdate();
		System.out.println(sql);
			i++;
		}
		//6.关闭资源
		statement.close();
		connection.close();
		System.out.println("success");
	}
}
