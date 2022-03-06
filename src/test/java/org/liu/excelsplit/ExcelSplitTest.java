package org.liu.excelsplit;

import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.WorkbookUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.platform.commons.util.StringUtils;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

public class ExcelSplitTest {
    static String fileName = "22年1月对账单汇总.xlsx";
    static int startRowIndex = 3;
    static DataFormatter dataFormatter = new DataFormatter();
    public static void main(String[] args) throws Exception {
        ExcelReader customer = ExcelUtil.getReader("classpath:" + fileName, 0);
        String prefix = "/tmp/" + System.currentTimeMillis();
        String filePath = saveFile(prefix, fileName, "classpath:" + fileName);

        List<List<Object>> customerData = customer.read(startRowIndex);
        List<String> customerNames = new ArrayList<>();
        for (int i = 0; i < customerData.size(); i++) {
            List<Object> customerLine = customerData.get(i);
            String customerName = customerLine.get(0).toString();
            if(StringUtils.isBlank(customerName)){
                break;
            }
            customerNames.add(customerName.trim());
        }
        List<String> res = copyAndMerge(prefix,customerNames,filePath);
        System.out.println(res);
    }
    private static List<String> copyAndMerge(String prefix, List<String> customerNames, String path) throws Exception{
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
    private static void merge(String tmpFilePath,String customerName,String newFilePath) throws IOException {
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
            if(i0++ < 2 || Objects.equals(cellValue,"客户名称") || StringUtils.isBlank(cellValue) || Objects.equals(cellValue,customerName)){
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
    private static String saveFile(String prefix, String fileName, String path) throws Exception{
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
    private static void copyRow(Workbook workbook, Sheet worksheet, int sourceRowNum, Sheet destinationSheet) {
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
