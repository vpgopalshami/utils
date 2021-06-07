package com.vpgs.utils.excel.util;


import com.vpgs.utils.excel.model.BasicForm;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class WriteExcelFileFromJavaObject {
    private static String[] headerColumns = {"First Name", "Last Name", "Date Of Birth", "Is Married", "Email Address", "Contact Number", "Aadhar Number", "Submitted Date"};

    public void writeExcelFile(List<BasicForm> basicFormList) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        CreationHelper creationHelper = workbook.getCreationHelper();
        Sheet sheet = workbook.createSheet("User_Data");

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        //headerFont.setFontHeight((short) 16);
        //headerFont.setColor(Font.COLOR_RED);

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(headerFont);

        Row headerRow = sheet.createRow(0);

        for(int i=0; i<headerColumns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headerColumns[i]);
            cellStyle.setWrapText(true);
            cell.setCellStyle(cellStyle);
        }

        CellStyle dataCellStyle = workbook.createCellStyle();
        dataCellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("YYYY-MM-dd"));

        int roeNum = 1;

        for(BasicForm basicForm : basicFormList) {

            Row dataRow = sheet.createRow(roeNum++);
            dataRow.createCell(0).setCellValue(basicForm.getFirstName());
            dataRow.createCell(1).setCellValue(basicForm.getLastName());
            dataRow.createCell(2).setCellValue(basicForm.getDob());
            dataRow.createCell(3).setCellValue(basicForm.isMarried());
            dataRow.createCell(4).setCellValue(basicForm.getEmailAddress());
            dataRow.createCell(5).setCellValue(basicForm.getContactNumber());
            dataRow.createCell(6).setCellValue(basicForm.getAadharNumber());
            dataRow.createCell(7).setCellValue(basicForm.getSubmittedDate());
        }

        FileOutputStream fileOutputStream = new FileOutputStream("D:\\Gopal\\Projects\\files\\BasicForm.xlsx");
        workbook.write(fileOutputStream);
        workbook.close();
    }


    public static void main(String[] args) {
        WriteExcelFileFromJavaObject writeExcelFileFromJavaObject = new WriteExcelFileFromJavaObject();
        BasicForm basicForm = new BasicForm();
        List<BasicForm> basicFormList = new ArrayList<>();
        basicForm.setFirstName("Gopalshami");
        basicForm.setLastName("V.P");
        basicForm.setDob("27/09/1982");
        basicForm.setEmailAddress("vpgopalshami@gmail.com");
        basicForm.setContactNumber("9840023568");
        basicForm.setSubmittedDate("06/06/2021");
        basicForm.setAadharNumber("123569774654455");
        basicFormList.add(basicForm);
        try {
            writeExcelFileFromJavaObject.writeExcelFile(basicFormList);
        } catch(Exception e) {
            e.printStackTrace();
        }
    }
}
