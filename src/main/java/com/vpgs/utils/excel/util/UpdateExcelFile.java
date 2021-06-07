package com.vpgs.utils.excel.util;

import com.vpgs.utils.excel.model.BasicForm;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class UpdateExcelFile {

    public void updateExcelFile(String filePath, BasicForm basicForm) {
        try {
            FileInputStream fileInputStream = new FileInputStream(new File(filePath));
            Workbook workbook = getWorkbook(fileInputStream,filePath);
            Sheet sheet = workbook.getSheet("User_Data");
            updateExcel(sheet, basicForm);
            updateFile(workbook, filePath);

        } catch(Exception e) {
            e.printStackTrace();
        }
    }

    private Workbook getWorkbook(FileInputStream fileInputStream ,String excelFilePath) throws IOException {
        Workbook workbook = null;
        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(fileInputStream);
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(fileInputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }
        return workbook;
    }

    private void updateExcel(Sheet sheet, BasicForm basicForm) {
        int lastRowNum = sheet.getLastRowNum();
        Row row = sheet.createRow(++lastRowNum);
        row.createCell(0).setCellValue(basicForm.getFirstName());
        row.createCell(1).setCellValue(basicForm.getLastName());
        row.createCell(2).setCellValue(basicForm.getDob());
        row.createCell(3).setCellValue(basicForm.isMarried());
        Cell cell = row.createCell(4);
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle.setWrapText(true);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(basicForm.getEmailAddress());
        //row.createCell(4).setCellValue(basicForm.getEmailAddress());
        row.createCell(5).setCellValue(basicForm.getContactNumber());
        row.createCell(6).setCellValue(basicForm.getAadharNumber());
        row.createCell(7).setCellValue(basicForm.getSubmittedDate());
    }

    private void updateFile(Workbook workbook, String filePath) throws Exception{
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        workbook.write(fileOutputStream);
        workbook.close();
    }

    public static void main(String[] args) {
        UpdateExcelFile updateExcelFile = new UpdateExcelFile();
        BasicForm basicForm = new BasicForm();
        basicForm.setFirstName("Sahashra Amma");
        basicForm.setLastName("Gopal");
        basicForm.setDob("01/08/2014");
        basicForm.setEmailAddress("gsahashra@gmail.com");
        basicForm.setContactNumber("9715263398");
        basicForm.setSubmittedDate("06/06/2021");
        basicForm.setAadharNumber("123569774654455");
        try {
            updateExcelFile.updateExcelFile("D:\\Gopal\\Projects\\files\\BasicForm.xlsx", basicForm);
        } catch(Exception e) {
            e.printStackTrace();
        }
    }
}
