package com.sandeep.springboot.util;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ByteArrayResource;

import com.sandeep.springboot.readinglist.model.PreallocationCriteriaDTO;




/**
 * 5 July, 2017
 * 
 * @author Trantor
 *
 */
public class XLSWriter {

	private static final Logger logger = LoggerFactory.getLogger(XLSWriter.class);

	/**
	 * Write List objects to an xlsx file
	 * 
	 * @param criteriaDTOList
	 *            - List of PreallocationCriteriaDTO objects
	 * @param columns
	 *            - List of columns need to be written in xlsx
	 * @param filename
	 *            - filename
	 * @throws IOException
	 */
	public static ByteArrayResource writeObjectToXLS(List<PreallocationCriteriaDTO> criteriaDTOList, String[] columns) throws IOException {

		logger.debug("Entering writeObjectToXLS method");
		logger.debug("Recieved columns list :-" + Arrays.toString((columns)));

		// Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();

		// Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Preallocation Data");

		// Object properties will be written on first row as column headings
		int rownum = 0;
		Row columnRow = sheet.createRow(rownum++);

		int columnCellnum = 0;

		logger.debug("start writing column names");
		// write column headers
		for (String column : columns) {

			Cell cell = columnRow.createCell(columnCellnum++);
			cell.setCellValue(column);
			cell.setCellStyle(getColumnHeaderStyle(workbook));
		}

		logger.debug("start writing object data");
		// write object data
		/*for (PreallocationCriteriaDTO criteriaDTO : criteriaDTOList) {
			Row dataRow = sheet.createRow(rownum++);

			writeObjectData(criteriaDTO, dataRow, workbook);

		}*/

		// fit data to columns
		for (int i = 0; i < 25; i++) {
			sheet.autoSizeColumn(i);
		}

		// write to output stream
		//FileOutputStream out;
		//out = new FileOutputStream(new File(filename));
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		workbook.write(baos);
		
		ByteArrayResource resource= new ByteArrayResource(baos.toByteArray());
		//out.close();
		return resource;
		/*logger.debug("done writing object data");

		logger.debug("Exit writeObjectToXLS method");*/

	}

	/**
	 * Font for column heading
	 * 
	 * @param workbook
	 * @return
	 */
	private static XSSFFont getColumnHeaderFont(XSSFWorkbook workbook) {

		// create font
		XSSFFont font = workbook.createFont();
		font.setFontHeightInPoints((short) 14);
		font.setFontName("Arial");
		font.setColor(IndexedColors.GREEN.getIndex());
		font.setBold(true);
		font.setItalic(false);
		return font;

	}

	/**
	 * Borders for column headings
	 * 
	 * @param workbook
	 * @return
	 */
	private static CellStyle getColumnHeaderStyle(XSSFWorkbook workbook) {

		// Create cell style
		CellStyle style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		// Setting font to style
		style.setFont(getColumnHeaderFont(workbook));
		return style;
	}

	/**
	 * STyle for cells containing data
	 * 
	 * @param workbook
	 * @return
	 */
	private static CellStyle getCellDataStyle(XSSFWorkbook workbook) {

		// Create cell style
		CellStyle style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_LEFT);
		style.setVerticalAlignment(CellStyle.VERTICAL_JUSTIFY);
		return style;
	}

	/**
	 * FOr each DTO, write data in a new row
	 * 
	 * @param criteriaDTO
	 * @param dataRow
	 * @param workbook
	 */
	private static void writeObjectData(PreallocationCriteriaDTO criteriaDTO, Row dataRow, XSSFWorkbook workbook) {/*

		// Find Owner
		Owner owner = criteriaDTO.getOwner();

		Cell ownerName = dataRow.createCell(0);
		ownerName.setCellValue(owner != null ? owner.getName() : "");
		ownerName.setCellStyle(getCellDataStyle(workbook));

		Cell ownerCode = dataRow.createCell(1);
		ownerCode.setCellValue(owner != null ? owner.getCode() : "");
		ownerCode.setCellStyle(getCellDataStyle(workbook));

		// Find ReportingGroup
		ReportingGroup reportingGroup = criteriaDTO.getReportingGroup();

		Cell reportingGroupName = dataRow.createCell(2);
		reportingGroupName.setCellValue(reportingGroup != null ? reportingGroup.getName() : "");
		reportingGroupName.setCellStyle(getCellDataStyle(workbook));

		Cell reportingGroupCode = dataRow.createCell(3);
		reportingGroupCode.setCellValue(reportingGroup != null ? reportingGroup.getCode() : "");
		reportingGroupCode.setCellStyle(getCellDataStyle(workbook));

		Cell category1 = dataRow.createCell(4);
		category1.setCellValue(criteriaDTO.getCategoryOne() != null ? criteriaDTO.getCategoryOne().getName() : "");
		category1.setCellStyle(getCellDataStyle(workbook));

		Cell category2 = dataRow.createCell(5);
		category2.setCellValue(criteriaDTO.getCategoryTwo() != null ? criteriaDTO.getCategoryTwo().getName() : "");
		category2.setCellStyle(getCellDataStyle(workbook));

		Cell format = dataRow.createCell(6);
		format.setCellValue(criteriaDTO.getFormat() != null ? criteriaDTO.getFormat().getName() : "");
		format.setCellStyle(getCellDataStyle(workbook));

		Cell subFormat = dataRow.createCell(7);
		subFormat.setCellValue(criteriaDTO.getSubFormat() != null ? criteriaDTO.getSubFormat().getName() : "");
		subFormat.setCellStyle(getCellDataStyle(workbook));

		Cell itemCode = dataRow.createCell(8);
		itemCode.setCellValue(criteriaDTO.getItem() != null ? criteriaDTO.getItem().getItemCode() : "");
		itemCode.setCellStyle(getCellDataStyle(workbook));

		// Find BilltoMaster
		BilltoMaster billToMaster = criteriaDTO.getBillto();

		Cell billTo = dataRow.createCell(9);
		billTo.setCellValue(billToMaster != null ? billToMaster.getBilltoName() : "");
		billTo.setCellStyle(getCellDataStyle(workbook));

		Cell billToNbr = dataRow.createCell(10);
		billToNbr.setCellValue(billToMaster != null ? billToMaster.getBilltoNbr().toPlainString() : "");
		billToNbr.setCellStyle(getCellDataStyle(workbook));

		Cell defaultPreallocQuantity = dataRow.createCell(11);
		defaultPreallocQuantity.setCellValue(criteriaDTO.getDefPreallocationQty());
		defaultPreallocQuantity.setCellStyle(getCellDataStyle(workbook));

		Cell usableMinPreallocQuantity = dataRow.createCell(12);
		usableMinPreallocQuantity.setCellValue(criteriaDTO.getUsableMinPreallocationQty());
		usableMinPreallocQuantity.setCellStyle(getCellDataStyle(workbook));

		Cell percentageIfFallsBelow = dataRow.createCell(13);
		percentageIfFallsBelow.setCellValue(criteriaDTO.getPcAvailableInv());
		percentageIfFallsBelow.setCellStyle(getCellDataStyle(workbook));

		Cell preRelease = dataRow.createCell(14);
		preRelease.setCellValue(criteriaDTO.getPreReleaseInd() == null ? "No" : criteriaDTO.getPreReleaseInd());
		preRelease.setCellStyle(getCellDataStyle(workbook));

		Cell overrideNYP = dataRow.createCell(15);
		overrideNYP.setCellValue(criteriaDTO.getOverrideNypInd() == null ? "No" : criteriaDTO.getOverrideNypInd());
		overrideNYP.setCellStyle(getCellDataStyle(workbook));
	*/}

}