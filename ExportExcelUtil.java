package com.spr.hcase.utils;

import com.spr.hcase.excele.CaseStatistics;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.ui.ModelMap;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * spr
 *
 * @Description :
 * @Author : JiangQi Luo
 * @Date : 2022/6/24 16:45
 */
@Slf4j
public class ExportExcelUtil<T> {

	public ExportExcelUtil(){}

	/**
	 由于不是封装好的，功能没有很强大，表头需要自己手动构建，并且不能和查出来的数据行错位
	 title 标题，head 表头 String[] headers = {菜品,图片,价格},
	 List<List<String>> dataCollection =  new ArrayList()<数据行.size>。
	 dataCollection .add(数据行)
	 */
	public void exportExcel(String title, String[] headers, Collection<T> dataset, HttpServletResponse response) {
		try {
			response.setContentType("application/x-download");
			response.setHeader("Content-Disposition",
					"attachment;filename=" + URLEncoder.encode(title, "UTF-8"));
			response.setHeader("fileName", URLEncoder.encode(title, "UTF-8"));
			response.setHeader("Access-Control-Expose-Headers","Content-Disposition");
			response.setHeader("Access-Control-Allow-Headers","Content-Type,fileName,Content-Disposition");
			response.flushBuffer();

			exportExcelXlsx(title, headers, dataset, response.getOutputStream(), "yyyy-MM-dd");


		} catch (IOException e1) {
			log.error("",e1);
		}

	}

	private void exportExcelXlsx(String title, String[] headers,
	                             Collection<T> dataset, OutputStream out, String pattern) {
		// 声明一个工作薄
		// XSSFWorkbook workbook = new XSSFWorkbook();
		SXSSFWorkbook workbook = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
		// 生成一个表格
		Sheet sheet = workbook.createSheet(title);
		// 设置表格默认列宽度为20个字节
		sheet.setDefaultColumnWidth(20);
		// 生成一个样式

		CellStyle style = workbook.createCellStyle();
		// 设置这些样式
		java.awt.Color color = new java.awt.Color(192, 192, 192);
		style.setFillForegroundColor(new XSSFColor(color).getIndex());
		style.setAlignment(HorizontalAlignment.CENTER);
		// 生成一个字体
		Font font = workbook.createFont();
		font.setBold(true);
		font.setFontHeightInPoints((short) 12);
		// 把字体应用到当前的样式
		style.setFont(font);
		// 生成并设置另一个样式
		CellStyle style2 = workbook.createCellStyle();
		java.awt.Color color2 = new java.awt.Color(255, 255, 0);
		style2.setFillForegroundColor(new XSSFColor(color2).getIndex());

		style2.setAlignment(HorizontalAlignment.CENTER);
		style2.setVerticalAlignment(VerticalAlignment.CENTER);
		// 生成另一个字体
		Font font2 = workbook.createFont();
		// font2.setBold(true);
		// 把字体应用到当前的样式
		style2.setFont(font2);

		// 声明一个画图的顶级管理器
		Drawing patriarch = sheet.createDrawingPatriarch();

		// 产生表格标题行
		Row row = sheet.createRow(0);
		for (int i = 0; i < headers.length; i++) {
			Cell cell = row.createCell(i);
			cell.setCellStyle(style);
			XSSFRichTextString text = new XSSFRichTextString(headers[i]);
			cell.setCellValue(text);
		}

		// 遍历集合数据，产生数据行
		Iterator<T> it = dataset.iterator();
		int index = 0;
		Cell cell = null;
		while (it.hasNext()) {
			index++;
			row = sheet.createRow(index);
			T t = it.next();
			// 利用反射，根据javabean属性的先后顺序，动态调用getXxx()方法得到属性值
			Field[] fields = t.getClass().getDeclaredFields();
			for (int i = 0; i < fields.length; i++) {
				if((i+1)>headers.length){
					break;
				}
				cell = row.createCell(i);
				cell.setCellStyle(style2);
				Field field = fields[i];
				String fieldName = field.getName();
				String getMethodName = "get"
						+ fieldName.substring(0, 1).toUpperCase()
						+ fieldName.substring(1);
				try {
					Class tCls = t.getClass();
					Method getMethod = tCls.getMethod(getMethodName);
					Object value = getMethod.invoke(t);
					// 判断值的类型后进行强制类型转换
					String textValue = "";
					if (value instanceof Boolean) {
						boolean bValue = (Boolean) value;
						textValue = "是";
						if (!bValue) {
							textValue = "否";
						}
					} else if (value instanceof Date) {
						Date date = (Date) value;
						SimpleDateFormat sdf = new SimpleDateFormat(pattern);
						textValue = sdf.format(date);
					} else if (value instanceof byte[]) {
						// 有图片时，设置行高为60px
						row.setHeightInPoints(60);
						// 设置图片所在列宽度为80px,注意这里单位的一个换算
						sheet.setColumnWidth(i, (short) (35.7 * 80));
						byte[] bsValue = (byte[]) value;
						ClientAnchor anchor = new XSSFClientAnchor(0, 0,
								1023, 255, (short) 6, index, (short) 6, index);
						anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
						patriarch.createPicture(anchor, workbook.addPicture(
								bsValue, XSSFWorkbook.PICTURE_TYPE_JPEG));
					} else {
						// 其它数据类型都当作字符串简单处理
						if (value != null) {
							textValue = value.toString();
						} else {
							textValue = "";
						}
					}
					// 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
					if (textValue != null) {
						Pattern p = Pattern.compile("^//d+(//.//d+)?$");
						Matcher matcher = p.matcher(textValue);
						if (matcher.matches()) {
							// 是数字当作double处理
							cell.setCellValue(Double.parseDouble(textValue));
						} else {
							RichTextString richString = new XSSFRichTextString(
									textValue);
							richString.applyFont(font2);
							cell.setCellValue(richString);
						}
					}
				} catch (SecurityException | IllegalAccessException | NoSuchMethodException | IllegalArgumentException | InvocationTargetException e) {
					log.error(e.getMessage(),e);
				}
			}
		}
		try {
			workbook.write(out);
		} catch (IOException e) {
			log.error(e.getMessage(),e);
		}
	}

	public void excel(){
		try{
			HSSFWorkbook  wb  =  new HSSFWorkbook();
			HSSFSheet sheet  =  wb.createSheet("11111");
			//设置表格的样式，可设置多个
			HSSFCellStyle style1 = wb.createCellStyle();
			style1.setVerticalAlignment(VerticalAlignment.CENTER);//水平
			style1.setAlignment(HorizontalAlignment.CENTER);//水平
			Font ztFont = wb.createFont();
			ztFont.setColor(Font.COLOR_NORMAL);
			ztFont.setFontHeightInPoints((short)16);
			ztFont.setFontName("宋体");
			style1.setFont(ztFont);

			HSSFCellStyle styleWork = wb.createCellStyle();
			styleWork.setVerticalAlignment(VerticalAlignment.CENTER);//水平
			styleWork.setAlignment(HorizontalAlignment.CENTER);//水平
			styleWork.setBorderBottom(BorderStyle.MEDIUM); //下边框
			styleWork.setBorderLeft(BorderStyle.MEDIUM);//左边框
			styleWork.setBorderTop(BorderStyle.MEDIUM);//上边框
			styleWork.setBorderRight(BorderStyle.MEDIUM);//右边框
			Font fontWork = wb.createFont();
			fontWork.setFontName("宋体");
			fontWork.setFontHeightInPoints((short)12);
			styleWork.setFont(fontWork);

			HSSFCellStyle styleValue = wb.createCellStyle();
			styleValue.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
			styleValue.setAlignment(HorizontalAlignment.CENTER);//水平
			Font fontValue = wb.createFont();
			fontValue.setFontName("宋体");
			fontValue.setFontHeightInPoints((short)12);
			styleValue.setFont(fontValue);

			//设置行
			HSSFRow row  =  sheet.createRow((short)   0);
			HSSFRow  row2  =  sheet.createRow((short)   2);
			HSSFRow  row3  =  sheet.createRow((short)   3);
			HSSFRow  row4  =  sheet.createRow((short)   4);

			HSSFCell ce=row.createCell((short)0);
			//这个是设置合并单元格的方法，四个参数分别是：起始行，结束行，起始列，结束列
			CellRangeAddress cra1 = new CellRangeAddress(0, 1, 0, 31);
			sheet.addMergedRegion(cra1);
			//给表格设置值
			ce.setCellValue("2021案件类型报表");
			ce.setCellStyle(style1);

			HSSFCell cellW = row2.createCell((short) (0));
			CellRangeAddress cra = new CellRangeAddress(3, 4, 0, 31);
			sheet.addMergedRegion(cra);
			cellW.setCellValue("委员会名称");
			cellW.setCellStyle(styleWork);

			HSSFCell cellWV = row2.createCell((short) (2));
			sheet.addMergedRegion(new  CellRangeAddress(2,3,2,3));
			cellWV.setCellValue("建设工成程");
			cellWV.setCellStyle(styleValue);

			HSSFCell cellPro = row3.createCell((short) (0));
			HSSFCell cellProVal = row3.createCell((short) (1));
			cellPro.setCellValue("人物2");
			sheet.addMergedRegion(new  CellRangeAddress(3,3,1,8));
			cellProVal.setCellValue("李四");
			cellPro.setCellStyle(styleWork);
			cellProVal.setCellStyle(styleValue);

			HSSFCell cellWorkFlow = row4.createCell((short) (0));
			HSSFCell cellWorkFlowVal = row4.createCell((short) (1));
			cellWorkFlow.setCellValue("人物3");
			sheet.addMergedRegion(new  CellRangeAddress(4,4,1,8));
			cellWorkFlowVal.setCellValue("王五");
			cellWorkFlow.setCellStyle(styleWork);
			cellWorkFlowVal.setCellStyle(styleValue);

//			//给合并的单元格设置边框，一定要统一放在下面，排查了好久的问题
//			CellRangeAddress cellRange = new CellRangeAddress(0,1,0,8);
//			sheet.addMergedRegion(cellRange);//为合并单元格添加边框
//			RegionUtil.setBorderTop(1, cellRange, sheet, wb);
//			RegionUtil.setBorderBottom(1, cellRange, sheet, wb);
//			RegionUtil.setBorderLeft(1, cellRange, sheet, wb);
//			RegionUtil.setBorderRight(1, cellRange, sheet, wb);
//
//			CellRangeAddress cellRange1 = new CellRangeAddress(2,2,1,8);
//			sheet.addMergedRegion(cellRange1);//为合并单元格添加边框
//			RegionUtil.setBorderTop(1, cellRange1, sheet, wb);
//			RegionUtil.setBorderBottom(1, cellRange1, sheet, wb);
//			RegionUtil.setBorderLeft(1, cellRange1, sheet, wb);
//			RegionUtil.setBorderRight(1, cellRange1, sheet, wb);
//
//			CellRangeAddress cellRange2 = new CellRangeAddress(3,3,1,8);
//			sheet.addMergedRegion(cellRange2);//为合并单元格添加边框
//			RegionUtil.setBorderTop(1, cellRange2, sheet, wb);
//			RegionUtil.setBorderBottom(1, cellRange2, sheet, wb);
//			RegionUtil.setBorderLeft(1, cellRange2, sheet, wb);
//			RegionUtil.setBorderRight(1, cellRange2, sheet, wb);
//
//			CellRangeAddress cellRange3 = new CellRangeAddress(4,4,1,8);
//			sheet.addMergedRegion(cellRange3);//为合并单元格添加边框
//			RegionUtil.setBorderTop(1, cellRange3, sheet, wb);
//			RegionUtil.setBorderBottom(1, cellRange3, sheet, wb);
//			RegionUtil.setBorderLeft(1, cellRange3, sheet, wb);
//			RegionUtil.setBorderRight(1, cellRange3, sheet, wb);
//
//			//循环设置表格宽度
//			for (int i = 0; i < 9; i++) {
//				sheet.autoSizeColumn(i);
//				sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 35 / 15);
//			}
//			//在第五行的基础上，创建单元格
//			HSSFCell  cell1  =   row5.createCell((short) (0) );
//			HSSFCell  cell2  =   row5.createCell((short) (1) );
//			HSSFCell  cell3  =   row5.createCell((short) (2) );
//			HSSFCell  cell4  =   row5.createCell((short) (3) );
//			HSSFCell  cellb1  =   row5.createCell((short) (4) );
//			HSSFCell  cellb2  =   row5.createCell((short) (5) );
//			HSSFCell  cellb3  =   row5.createCell((short) (6) );
//			HSSFCell  cellb4  =   row5.createCell((short) (7) );
//			HSSFCell  cellb5  =   row5.createCell((short) (8) );
//			cell1.setCellValue("test1");
//			cell2.setCellValue("test2");
//			cell3.setCellValue("test3");
//			cell4.setCellValue("test4");
//			cellb1.setCellValue("test5");
//			cellb2.setCellValue("test6");
//			cellb3.setCellValue("test7");
//			cellb4.setCellValue("test8");
//			cellb5.setCellValue("test9");
//
//			cell1.setCellStyle(styleWork);
//			cell2.setCellStyle(styleWork);
//			cell3.setCellStyle(styleWork);
//			cell4.setCellStyle(styleWork);
//			cellb1.setCellStyle(styleWork);
//			cellb2.setCellStyle(styleWork);
//			cellb3.setCellStyle(styleWork);
//			cellb4.setCellStyle(styleWork);
//			cellb5.setCellStyle(styleWork);
			FileOutputStream fileOut  =  new  FileOutputStream("D:测试表格.xls");
			wb.write(fileOut);
			fileOut.close();
		}catch(Exception ex){
			ex.printStackTrace();
		}
	}



}
