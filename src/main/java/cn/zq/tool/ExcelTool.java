package cn.zq.tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import cn.zq.model.ResponseData;

/**
 * excel工具
 */

public class ExcelTool {

	/**
	 * 判断列格式是否正确，若格式正确，返回列值的字符形式；若格式错误，则抛出异常
	 * 
	 * @param Cell
	 *            excel列对象
	 */
	protected static String getCellValue(Cell cell) throws Exception {
		if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			return cell.getStringCellValue().trim();
		}
		if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
			return cell.getStringCellValue().trim();
		}
		return "";
	}

	/**
	 * excel日期转化为java日期
	 * 
	 * @param days
	 *            excel日期的数值
	 * @return
	 */
	protected Date getExcelDate(int days) {
		Calendar c3 = Calendar.getInstance();
		c3.set(1900, 0, -1);
		c3.add(Calendar.DATE, days);
		Date date = c3.getTime();
		return date;
	}

	/**
	 * excel列名
	 * 
	 * @param index
	 *            excel列序号
	 * @return excel列名，如“A，AB”
	 */
	protected static String getExcelColName(int index) {
		if (index <= 26)
			return String.valueOf((char) ('A' + index - 1));
		int a = index / 26;
		int b = index % 26;

		return String.valueOf((char) ('A' + a - 1)) + String.valueOf((char) ('A' + b - 1));

	}

	/**
	 * 判断导入数据行是否为空
	 *
	 * @param row
	 *            excel行对象
	 */
	protected Boolean isRowBlank(Row row) {
		if ((row.getCell(0) == null || row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_BLANK)
				&& (row.getCell(3) == null || row.getCell(3).getCellType() == HSSFCell.CELL_TYPE_BLANK)
				&& (row.getCell(4) == null || row.getCell(4).getCellType() == HSSFCell.CELL_TYPE_BLANK)
				&& (row.getCell(5) == null || row.getCell(5).getCellType() == HSSFCell.CELL_TYPE_BLANK)
				&& (row.getCell(7) == null || row.getCell(7).getCellType() == HSSFCell.CELL_TYPE_BLANK)
				&& (row.getCell(9) == null || row.getCell(9).getCellType() == HSSFCell.CELL_TYPE_BLANK)
				&& (row.getCell(10) == null || row.getCell(10).getCellType() == HSSFCell.CELL_TYPE_BLANK)
				&& (row.getCell(11) == null || row.getCell(11).getCellType() == HSSFCell.CELL_TYPE_BLANK)
				&& (row.getCell(12) == null || row.getCell(12).getCellType() == HSSFCell.CELL_TYPE_BLANK)
				&& (row.getCell(13) == null || row.getCell(13).getCellType() == HSSFCell.CELL_TYPE_BLANK))
			return true;
		return false;
	}

	/**
	 * 根据出生日期计算年龄
	 * 
	 * @param days
	 *            excel日期的数值
	 * @return
	 */
	protected int getAge(Date birthDay) throws Exception {
		if (birthDay == null)
			return 0;
		// 获取当前系统时间
		Calendar cal = Calendar.getInstance();
		// 如果出生日期大于当前时间，则抛出异常
		if (cal.before(birthDay)) {
			throw new IllegalArgumentException("The birthDay is before Now.It's unbelievable!");
		}
		// 取出系统当前时间的年、月、日部分
		int yearNow = cal.get(Calendar.YEAR);
		int monthNow = cal.get(Calendar.MONTH);
		int dayOfMonthNow = cal.get(Calendar.DAY_OF_MONTH);

		// 将日期设置为出生日期
		cal.setTime(birthDay);
		// 取出出生日期的年、月、日部分
		int yearBirth = cal.get(Calendar.YEAR);
		int monthBirth = cal.get(Calendar.MONTH);
		int dayOfMonthBirth = cal.get(Calendar.DAY_OF_MONTH);
		// 当前年份与出生年份相减，初步计算年龄
		int age = yearNow - yearBirth;
		// 当前月份与出生日期的月份相比，如果月份小于出生月份，则年龄上减1，表示不满多少周岁
		if (monthNow <= monthBirth) {
			// 如果月份相等，在比较日期，如果当前日，小于出生日，也减1，表示不满多少周岁
			if (monthNow == monthBirth) {
				if (dayOfMonthNow < dayOfMonthBirth)
					age--;
			} else {
				age--;
			}
		}
		return age;
	}

	/**
	 * 判断列格式是否正确，若格式正确，返回列值；若格式错误，则返回对应的提示信息
	 * 
	 * @param Cell
	 *            excel列对象
	 * @param rowIndex
	 *            第几行
	 * @param colIndex
	 *            第几列
	 */
	private static ResponseData<String> getCellValue(Cell cell, int rowIndex, int colIndex) {
		ResponseData<String> data = new ResponseData<String>();
		if (cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
			data.setData("");
			data.setMessage("行号：" + rowIndex + ", 列号：" + getExcelColName(colIndex) + ", 数据为空");
			return data;
		}
		String value = null;
		switch (colIndex) {
		case 1:
			try {
				value = getCellValue(cell);
			} catch (Exception e) {
				data.setCode(-1);
				data.setMessage("行号：" + rowIndex + ", 列号：" + getExcelColName(colIndex) + ",  所属单位格式不对");
				return data;
			}
			break;

		case 3:// 收件人地址
			try {
				value = getCellValue(cell);
			} catch (Exception e) {
				data.setCode(-1);
				data.setMessage("行号：" + rowIndex + ", 列号：" + getExcelColName(colIndex) + ",  收件人地址格式不对");
				return data;
			}
			break;
		default:
			try {
				value = getCellValue(cell);
			} catch (Exception e) {
				data.setCode(-1);
				data.setMessage("行号：" + rowIndex + ", 列号：" + getExcelColName(colIndex) + ",  格式不对");
				return data;
			}
			break;
		}
		data.setData(value);
		data.setMessage("获取列的值成功！");
		return data;
	}

	/**
	 * excel文件获公司名称
	 * 
	 * @param fileName
	 *            excel文件名(包含完整路径)
	 * 
	 * @return
	 */
	public static void processData(String fileName) {

		ResponseData<String> cellObject = new ResponseData<String>(); // 列值对象
		String cellValue = null; // 列值
		HSSFWorkbook wb = null;
		List<String> list = new ArrayList<String>();
		List<String> list2 = new ArrayList<String>();
		try {
			//InputStream input = new FileInputStream(new File("C:\\Users\\darcy\\Desktop\\1.xls"));
			InputStream input = new FileInputStream(new File(fileName));
			wb = new HSSFWorkbook(input);
			HSSFSheet sheet = wb.getSheetAt(0);
			int rowIndex = 0;// 行号
			for (Iterator<Row> iter = (Iterator<Row>) sheet.rowIterator(); iter.hasNext();) {
				Row row = iter.next();
				rowIndex++;
				// 第一行是表头，非数据，跳过
				if (rowIndex == 1) {
					continue;
				}

				int cellCount = 3;
				int colIndex = 0;

				// begin of row
				for (int i = 0; i < cellCount; i++) {
					Cell cell = row.getCell(i);
					colIndex++;
					cellObject = getCellValue(cell, rowIndex, colIndex);
					cellValue = cellObject.getData();
					switch (colIndex) {
					case 1:
						list.add(cellValue);
						break;
					case 2:
						break;
					case 3:
						list2.add(cellValue);
						break;

					default:
						break;
					}
				}
			}

			// 写数据
			rowIndex = 0;// 行号
			for (Iterator<Row> iter = (Iterator<Row>) sheet.rowIterator(); iter.hasNext();) {
				Row row = iter.next();
				rowIndex++;
				// 第一行是表头，非数据，跳过
				if (rowIndex == 1) {
					continue;
				}
				boolean flag = isExist(row.getCell(0).getStringCellValue(), list2);
				Cell cell = row.createCell(1);
				cell.setCellValue(flag);
			}

			// 将excel写入
			OutputStream os = new FileOutputStream(fileName);
			wb.write(os);
			os.close();

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				wb.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	/**
	 * 判断公司名称是否在第三列里面存在
	 */
	private static boolean isExist(String name, List<String> list) {
		for (String a : list) {
			if (name.equals(a))
				return true;
		}
		return false;
	}

}
