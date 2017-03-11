package attpackage;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.*;

/**
 * Created by Dhruv Sb on 14-10-2016.
 */
public class OnlyAttendace {
    static void OnlyAttendanceMethod() {
        HSSFWorkbook workbookOut = new HSSFWorkbook();
        HSSFSheet sheetOut = workbookOut.createSheet("Sheet");
        FileInputStream FileInput = null;
        try {
            FileInput = new FileInputStream(new File("C:\\CO-III.xls"));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        HSSFWorkbook workbook = null;
        try {
            workbook = new HSSFWorkbook(FileInput);
        } catch (IOException e) {
            e.printStackTrace();
        }
        HSSFSheet sheetIn = workbook.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;
        CellStyle styleA = workbookOut.createCellStyle();
        styleA.setAlignment(CellStyle.ALIGN_CENTER);
        sheetOut.setColumnWidth(0,3500);
        sheetOut.setColumnWidth(1,5000);
        sheetOut.setColumnWidth(2,3000);
        row = sheetOut.createRow(0);
        cell = row.createCell(0);
        cell.getRow().setHeightInPoints(20);
        cell.setCellValue("Enrollment no.");
        cell = row.createCell(1);
        cell.setCellValue("Student's name");
        cell.setCellStyle(styleA);
        cell = row.createCell(2);
        cell.setCellValue("Percentage");
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
            row = sheetOut.createRow(i + 1);
            cell = row.createCell(0);
            cell.setCellValue(enroll[i]);
            cell.getRow().setHeightInPoints(20);
            cell.setCellStyle(styleA);
            row = sheetOut.getRow(i + 1);
            cell = row.createCell(1);
            cell.setCellValue(name[i]);
            cell.getRow().setHeightInPoints(20);
            cell.setCellStyle(styleA);
        }
        for (int k = 1; k < 5; ++k) {
            HSSFSheet sheet = workbook.getSheetAt(k);
            for (int i = 1; i < 11; ++i) {
                HSSFRow rowIn = sheet.getRow(i);
                if (rowIn.getCell(2).getCellType() == HSSFCell.CELL_TYPE_STRING) {
                    String absent = rowIn.getCell(2).getStringCellValue();
                    String[] absentNo = absent.split(",");
                    int total = absentNo.length;
                    for (int j = 0; j < total; ++j) {
                        Integer absentInt = Integer.valueOf(absentNo[j]);
                        ++AttendanceOverall[absentInt];
                    }
                    ++AttendanceOverall[0];
                } else if (rowIn.getCell(2).getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                    if (rowIn.getCell(2).getNumericCellValue() == 0) {
                        ++AttendanceOverall[0];
                    } else {
                        int absent1 = (int) rowIn.getCell(2).getNumericCellValue();
                        ++AttendanceOverall[absent1];
                        ++AttendanceOverall[0];
                    }
                }
            }
        }
        for (int i = 1; i < 51; ++i) {
            row = sheetOut.getRow(i);
            cell = row.createCell(2);
            double present = (double) AttendanceOverall[i];
            double total = (double) AttendanceOverall[0];
            double percentage = 100 * (1 - (present / total));
            cell.setCellValue(percentage);
            CellStyle style = workbookOut.createCellStyle();
            style.setAlignment(CellStyle.ALIGN_CENTER);
            if (percentage <= 50) {
                style.setFillForegroundColor(IndexedColors.RED.getIndex());
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                cell.setCellStyle(style);
            } else if (percentage <= 75) {
                style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                cell.setCellStyle(style);
            } else if (percentage > 95) {
                style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                cell.setCellStyle(style);
            }
            else cell.setCellStyle(style);
        }
        try {
            FileOutputStream output = new FileOutputStream("OnlyAttendance.xls");
            workbookOut.write(output);
            output.close();
        } catch (Exception exc) {
            exc.printStackTrace();
        }
    }
}