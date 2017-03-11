package attpackage;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;

/**
 * Created by Dhruv Sb on 13-10-2016.
 */
public class FullReport {
    public static void ReportMethod() throws IOException {
        HSSFWorkbook workbookOut = new HSSFWorkbook();
        HSSFSheet sheetOut = workbookOut.createSheet("Sheet");
        FileInputStream FileInput = new FileInputStream(new File("C:\\CO-I.xls"));
        HSSFWorkbook workbookIn = new HSSFWorkbook(FileInput);
        HSSFSheet sheetIn = workbookIn.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;
        sheetOut.setColumnWidth(0,3200);
        sheetOut.setColumnWidth(1,5000);
        sheetOut.setColumnWidth(10,4500);
        CellStyle styleA = workbookOut.createCellStyle();
        styleA.setAlignment(CellStyle.ALIGN_CENTER);
        row = sheetOut.createRow(1);
        cell = row.createCell(0);
        cell.setCellValue("Enrollment no.");
        cell.setCellStyle(styleA);
        cell.getRow().setHeightInPoints(20);
        cell = row.createCell(1);
        cell.setCellStyle(styleA);
        cell.setCellValue("Name");
        for (int i = 1; i < 5; ++i) {
            cell = row.createCell(i * 2);
            cell.setCellValue("Lectures");
            cell.setCellStyle(styleA);
            cell = row.createCell((i * 2)+1);
            cell.setCellValue("%");
            sheetOut.addMergedRegion(new CellRangeAddress(0,0,i * 2,(i * 2) + 1));
            cell.setCellStyle(styleA);
        }
        cell = row.createCell(10);
        cell.setCellValue("Overall Attendance");
        cell.setCellStyle(styleA);
        String[] name = new String[50];
        int[] enroll = new int[50];
        for (int i = 0; i < 50; ++i) {
            row = sheetIn.getRow(i + 1);
            int enrollIn = (int) row.getCell(0).getNumericCellValue();
            enroll[i] = enrollIn;
            row = sheetIn.getRow(i + 1);
            String nameIn = row.getCell(1).getStringCellValue();
            name[i] = nameIn;
        }
        int[] AttendanceOverall = new int[51];
        for (int i = 0; i < 51; ++i)
            AttendanceOverall[i] = 0;
        for (int i = 0; i < 50; ++i) {
            row = sheetOut.createRow(i + 2);
            cell = row.createCell(0);
            cell.getRow().setHeightInPoints(20);
            cell.setCellValue(enroll[i]);
            cell.setCellStyle(styleA);
            row = sheetOut.getRow(i + 2);
            cell = row.createCell(1);
            cell.setCellValue(name[i]);
            cell.setCellStyle(styleA);
        }
        HSSFRow row1 = sheetOut.createRow(0);
        for (int k = 1; k < 5; ++k) {
            String SubjectName = workbookIn.getSheetName(k);
            cell = row1.createCell(2*k);
            cell.setCellValue(SubjectName);
            cell.getRow().setHeightInPoints(20);
            cell.setCellStyle(styleA);
            HSSFSheet sheet = workbookIn.getSheetAt(k);
            int[] SubjectAttendance = new int[51];
            for (int i = 0; i < 51; ++i)
                SubjectAttendance[i] = 0;
            for (int i = 1; i < 11; ++i) {
                HSSFRow rowIn = sheet.getRow(i);
                if (rowIn.getCell(2).getCellType() == HSSFCell.CELL_TYPE_STRING) {
                    String absent = rowIn.getCell(2).getStringCellValue();
                    String[] absentNo = absent.split(",");
                    int total = absentNo.length;
                    for (int j = 0; j < total; ++j) {
                        Integer absentInt = Integer.valueOf(absentNo[j]);
                        ++SubjectAttendance[absentInt];
                        ++AttendanceOverall[absentInt];
                    }
                    ++SubjectAttendance[0];
                    ++AttendanceOverall[0];
                } else if (rowIn.getCell(2).getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                    if (rowIn.getCell(2).getNumericCellValue() == 0) {
                        ++SubjectAttendance[0];
                        ++AttendanceOverall[0];
                    } else {
                        int absent1 = (int) rowIn.getCell(2).getNumericCellValue();
                        ++SubjectAttendance[absent1];
                        ++AttendanceOverall[absent1];
                        ++SubjectAttendance[0];
                        ++AttendanceOverall[0];
                    }
                }
            }
            for (int i = 0; i < 50; ++i) {
                double present = (double) SubjectAttendance[i + 1];
                double total = (double) SubjectAttendance[0];
                double percentage = 100 * (1 - (present / total));
                row = sheetOut.getRow(i + 2);
                cell = row.createCell((k * 2));
                String SubAttended = (SubjectAttendance[0] - SubjectAttendance[i + 1]) + "/" + SubjectAttendance[0];
                cell.setCellValue(SubAttended);
                cell.setCellStyle(styleA);
                cell = row.createCell((k * 2) + 1);
                cell.setCellValue(percentage);
                cell.setCellStyle(styleA);
                CellStyle style = workbookOut.createCellStyle();
                if (percentage < 50) {
                    style.setFillForegroundColor(IndexedColors.RED.getIndex());
                    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                    style.setAlignment(CellStyle.ALIGN_CENTER);
                    cell.setCellStyle(style);
                }
            }
        }
        for (int i = 1; i < 51; ++i) {
            row = sheetOut.getRow(i+1);
            cell = row.createCell(10);
            double present = (double) AttendanceOverall[i];
            double total = (double) AttendanceOverall[0];
            double percentage = 100 * (1 - (present / total));
            cell.setCellValue(percentage);
            cell.setCellStyle(styleA);
            CellStyle style = workbookOut.createCellStyle();
            if (percentage <= 50) {
                style.setFillForegroundColor(IndexedColors.RED.getIndex());
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                style.setBorderBottom(CellStyle.BORDER_THIN);
                style.setBottomBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setBorderLeft(CellStyle.BORDER_THIN);
                style.setLeftBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setBorderRight(CellStyle.BORDER_THIN);
                style.setRightBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setAlignment(CellStyle.ALIGN_CENTER);
                cell.setCellStyle(style);
            } else if (percentage <= 75) {
                style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                style.setBorderLeft(CellStyle.BORDER_THIN);
                style.setLeftBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setBorderRight(CellStyle.BORDER_THIN);
                style.setRightBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setBorderBottom(CellStyle.BORDER_THIN);
                style.setBottomBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setAlignment(CellStyle.ALIGN_CENTER);
                cell.setCellStyle(style);
            } else if (percentage > 95) {
                style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                style.setBorderBottom(CellStyle.BORDER_THIN);
                style.setBottomBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setBorderLeft(CellStyle.BORDER_THIN);
                style.setLeftBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setBorderRight(CellStyle.BORDER_THIN);
                style.setRightBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setAlignment(CellStyle.ALIGN_CENTER);
                cell.setCellStyle(style);
            }else{
                style.setBorderBottom(CellStyle.BORDER_THIN);
                style.setBottomBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setBorderLeft(CellStyle.BORDER_THIN);
                style.setLeftBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setBorderRight(CellStyle.BORDER_THIN);
                style.setRightBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
                style.setAlignment(CellStyle.ALIGN_CENTER);
                cell.setCellStyle(style);
            }
            row = sheetOut.createRow(52);
            cell = row.createCell(1);
            cell.setCellValue("Average attendance:");
            cell.setCellStyle(styleA);
            cell.getRow().setHeightInPoints(20);
            cell = row.createCell(3);
            cell.setCellFormula("AVERAGE(D3:D52)");
            cell.setCellStyle(styleA);
            cell = row.createCell(5);
            cell.setCellFormula("AVERAGE(F3:F52)");
            cell.setCellStyle(styleA);
            cell = row.createCell(7);
            cell.setCellFormula("AVERAGE(H3:H52)");
            cell.setCellStyle(styleA);
            cell = row.createCell(9);
            cell.setCellFormula("AVERAGE(J3:J52)");
            cell.setCellStyle(styleA);
            cell = row.createCell(10);
            cell.setCellFormula("AVERAGE(K3:K52)");
            cell.setCellStyle(styleA);
        }
        try {
            FileOutputStream FileOutput = new FileOutputStream("FullReport.xls");
            workbookOut.write(FileOutput);
            FileOutput.close();
        } catch (Exception exc) {
            exc.printStackTrace();
        }
    }
}