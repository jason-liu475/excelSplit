package org.liu.excelsplit.controller;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.WorkbookUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Controller;
import org.springframework.util.ObjectUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.*;

@Slf4j
@Controller
public class ExcelSplitController {
	private DataFormatter dataFormatter = new DataFormatter();
	@ResponseBody
	@PostMapping(value="/upload")
	public List<String> upload(@RequestParam("file")MultipartFile[] files) throws Exception {
		List<String> filePaths = new ArrayList<>();
		if(ArrayUtil.isNotEmpty(files)){
			String prefix = "/tmp/" + System.currentTimeMillis();
			for(MultipartFile file : files){
				String fileName = file.getOriginalFilename();
				assert fileName != null;
				if(!fileName.contains("xls") && !fileName.contains("xlsx")){
					continue;
				}
				log.info("file name:{}",fileName);
				String filePath = saveFile(prefix,fileName,file);
				ExcelReader customer = ExcelUtil.getReader(filePath, 0);

				List<List<Object>> customerData = customer.read(2);
				customer.close();
				List<String> customerNames = new ArrayList<>();
				for (int i = 0; i < customerData.size(); i++) {
					List<Object> customerLine = customerData.get(i);
					String customerName = customerLine.get(0).toString();
					if(ObjectUtils.isEmpty(customerName)){
						break;
					}
					customerNames.add(customerName.trim());
				}
				return copyAndMerge(prefix,customerNames,filePath);
			}
		}
		return filePaths;
	}

	@RequestMapping(value="/download",method = RequestMethod.GET)
	public void download(@RequestParam("filePath")String filePath,HttpServletResponse response) throws Exception {
		String fileName = filePath.substring(filePath.lastIndexOf("/") + 1);
		response.reset();
		response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
		response.setHeader("Connection", "close");
		response.setHeader("Content-Type", "application/octet-stream");
		writeResponse(filePath,response);
	}
	private String saveFile(String prefix,String fileName,MultipartFile file) throws Exception{
		File dir = new File(prefix);
		if(!dir.exists()){
			dir.mkdirs();
		}
		InputStream inputStream = file.getInputStream();
		String filePath = prefix + "/" + fileName;
		try(OutputStream out = new FileOutputStream(filePath)){
			byte[] bs = new byte[1024];
			int len;
			while ((len = inputStream.read(bs)) != -1) {
				out.write(bs, 0, len);
			}
		}
		return filePath;
	}
	private void writeResponse(String filePath,HttpServletResponse response) throws Exception{
		ServletOutputStream outputStream = response.getOutputStream();
		FileInputStream fis = new FileInputStream(filePath);
		int len;
		byte[] bs = new byte[1024];
		while ((len = fis.read(bs)) != -1) {
			outputStream.write(bs,0,len);
		}
		fis.close();
		outputStream.flush();
		outputStream.close();
	}
	private List<String> copyAndMerge(String prefix, List<String> customerNames, String path) throws Exception{
		List<String> res = new ArrayList<>();
		for (int i = 0; i < customerNames.size(); i++) {
			String customerName = customerNames.get(i);
			String tmpFilePath = saveFile(prefix, System.currentTimeMillis() + ".xlsx", path);
			String newFilePath = prefix + "/" + customerName + ".xlsx";
			merge(tmpFilePath,customerName,newFilePath);
			res.add(newFilePath);
		}
		return res;
	}
	private void merge(String tmpFilePath,String customerName,String newFilePath) throws IOException {
		File tmpFile = FileUtil.file(tmpFilePath);
		Workbook bookForWriter = WorkbookUtil.createBook(tmpFile);
		Sheet sheet0 = bookForWriter.getSheetAt(0);
		String sheet0Name = sheet0.getSheetName();
		bookForWriter.setSheetName(0,String.valueOf(System.currentTimeMillis()));
		Sheet sheet2 = bookForWriter.createSheet(sheet0Name);
		Iterator<Row> row0Iterator = sheet0.rowIterator();
		int i0 = 0;
		while(row0Iterator.hasNext()) {
			Row row = row0Iterator.next();
			String cellValue = dataFormatter.formatCellValue(row.getCell(0)).trim();
			if(i0++ < 2 || Objects.equals(cellValue,"客户名称") || ObjectUtils.isEmpty(cellValue) || Objects.equals(cellValue,customerName)){
				copyRow(bookForWriter,sheet0,row.getRowNum(),sheet2);
			}
		}

		Sheet sheet1 = bookForWriter.getSheetAt(1);
		String sheet1Name = sheet1.getSheetName();
		bookForWriter.setSheetName(1,String.valueOf(System.currentTimeMillis())+0);
		Sheet sheet3 = bookForWriter.createSheet(sheet1Name);
		Iterator<Row> row1Iterator = sheet1.rowIterator();

		while(row1Iterator.hasNext()) {
			Row row = row1Iterator.next();
			String cellValue = dataFormatter.formatCellValue(row.getCell(0)).trim();
			if(Objects.equals(cellValue,"客户名称") || Objects.equals(cellValue,customerName)){
				copyRow(bookForWriter,sheet1,row.getRowNum(),sheet3);
			}
		}
		bookForWriter.removeSheetAt(0);
		bookForWriter.removeSheetAt(0);
		BufferedOutputStream outputStream = FileUtil.getOutputStream(FileUtil.file(newFilePath));
		bookForWriter.write(outputStream);
		outputStream.close();
		bookForWriter.close();
		tmpFile.delete();
	}
	private String saveFile(String prefix, String fileName, String path) throws Exception{
		File dir = new File(prefix);
		if(!dir.exists()){
			dir.mkdirs();
		}
		InputStream inputStream = FileUtil.getInputStream(path);
		String filePath = prefix + "/" + fileName;
		try(OutputStream out = new FileOutputStream(filePath)){
			byte[] bs = new byte[1024];
			int len;
			while ((len = inputStream.read(bs)) != -1) {
				out.write(bs, 0, len);
			}
		}
		inputStream.close();
		return filePath;
	}
	private void copyRow(Workbook workbook, Sheet worksheet, int sourceRowNum, Sheet destinationSheet) {
		// Get the source / new row
		int destinationRowNum = destinationSheet.getLastRowNum() + 1;
		Row newRow = destinationSheet.createRow(destinationRowNum);
		Row sourceRow = worksheet.getRow(sourceRowNum);

		// Loop through source columns to add to new row
		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			// Grab a copy of the old/new cell
			Cell oldCell = sourceRow.getCell(i);
			Cell newCell = newRow.createCell(i);

			// If the old cell is null jump to next cell
			if (oldCell == null) {
				newCell = null;
				continue;
			}

			// Copy style from old cell and apply to new cell
			CellStyle newCellStyle = workbook.createCellStyle();
			newCellStyle.cloneStyleFrom(oldCell.getCellStyle());

			newCell.setCellStyle(newCellStyle);

			// If there is a cell comment, copy
			if (oldCell.getCellComment() != null) {
				newCell.setCellComment(oldCell.getCellComment());
			}

			// If there is a cell hyperlink, copy
			if (oldCell.getHyperlink() != null) {
				newCell.setHyperlink(oldCell.getHyperlink());
			}

			// Set the cell data type
//            newCell.setCellType(oldCell.getCellType());

			// Set the cell data value
			switch (oldCell.getCellType()) {
				case BLANK:
					newCell.setCellValue(oldCell.getStringCellValue());
					break;
				case BOOLEAN:
					newCell.setCellValue(oldCell.getBooleanCellValue());
					break;
				case ERROR:
					newCell.setCellErrorValue(oldCell.getErrorCellValue());
					break;
				case FORMULA:
					newCell.setCellFormula(oldCell.getCellFormula());
					break;
				case NUMERIC:
					newCell.setCellValue(oldCell.getNumericCellValue());
					break;
				case STRING:
					newCell.setCellValue(oldCell.getRichStringCellValue());
					break;
			}
		}

		// If there are are any merged regions in the source row, copy to new row
		for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
			CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
			if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
				CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
						(newRow.getRowNum() +
								(cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
								)),
						cellRangeAddress.getFirstColumn(),
						cellRangeAddress.getLastColumn());
				destinationSheet.addMergedRegion(newCellRangeAddress);
			}
		}
	}
}
