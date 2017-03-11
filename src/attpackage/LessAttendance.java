package attpackage;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Dhruv Sb on 14-10-2016.
 */
public class LessAttendance {
    public static void LAmethod() throws IOException {
        HSSFWorkbook workbookOut = new HSSFWorkbook();
        HSSFSheet sheetOut = workbookOut.createSheet("Sheet");
        FileInputStream FileInput = new FileInputStream(new File("C:\\CO-I.xls"));
        HSSFWorkbook workbook = new HSSFWorkbook(FileInput);
        HSSFSheet sheetIn = workbook.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;
        CellStyle styleA = workbookOut.createCellStyle();
        styleA.setAlignment(CellStyle.ALIGN_CENTER);
        sheetOut.setColumnWidth(0,3500);
        sheetOut.setColumnWidth(1,5000);
        sheetOut.setColumnWidth(2,3000);
        int[] AttendanceOverall = new int[51];
        row = sheetOut.createRow(0);
        cell = row.createCell(0);
        cell.getRow().setHeightInPoints(20);
        cell.setCellStyle(styleA);
        cell.setCellValue("Enrollment no.");
        cell = row.createCell(1);
        cell.setCellValue("Student's name");
        cell.setCellStyle(styleA);
        cell = row.createCell(2);
        cell.setCellValue("Percentage");
        cell.setCellStyle(styleA);
        int LAcount = 0;
        for (int i = 0; i < 51; ++i)
            AttendanceOverall[i] = 0;
        for (int k = 1; k < 5; ++k) {
            HSSFSheet sheet = workbook.getSheetAt(k);
            for (int i = 1; i < 11; ++i) {
                HSSFRow rowIn = sheet.getRow(i);
                if (rowIn.getCell(2).getCellType() == HSSFCell.CELL_TYPE_STRING) {
                    String absent = rowIn.getCell(2).getStringCellValue();
                    String[] absentNo = absent.split(",");
                    int size = absentNo.length;
                    for (int j = 0; j < size; ++j) {
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
        double[] attendanceLA = new double[50];
        for (int i = 1; i < 51; ++i) {
            double present = (double) AttendanceOverall[i];
            double total = (double) AttendanceOverall[0];
            double percentage = 100 * (1 - (present / total));
            attendanceLA[i-1] = percentage;
            if (percentage < 50)
                ++LAcount;
        }
        String[] name = new String[LAcount];
        int[] enroll = new int[LAcount];
        double[] percentLA = new double[LAcount];
        int j = 0;
        for (int i = 0; i < 50; ++i) {
            if ( attendanceLA[i] < 50 ) {
                row = sheetIn.getRow(i + 1);
                int enrollIn = (int) row.getCell(0).getNumericCellValue();
                enroll[j] = enrollIn;
                row = sheetIn.getRow(i + 1);
                String nameIn = row.getCell(1).getStringCellValue();
                name[j] = nameIn;
                percentLA[j] = attendanceLA[i];
                ++j;
            }
        }
        for (int i = 0; i < LAcount; ++i) {
            row = sheetOut.createRow(i + 1);
            cell = row.createCell(0);
            cell.getRow().setHeightInPoints(20);
            cell.setCellValue(enroll[i]);
            cell.setCellStyle(styleA);
            row = sheetOut.getRow(i + 1);
            cell = row.createCell(1);
            cell.setCellStyle(styleA);
            cell.getRow().setHeightInPoints(20);
            cell.setCellValue(name[i]);
            row = sheetOut.getRow(i + 1);
            cell = row.createCell(2);
            cell.setCellStyle(styleA);
            cell.getRow().setHeightInPoints(20);
            cell.setCellValue(percentLA[i]);
        }
        try {
            FileOutputStream OutputFile = new FileOutputStream("LA.xls");
            workbookOut.write(OutputFile);
            OutputFile.close();
        } catch (Exception exc) {
            exc.printStackTrace();
        }
    }
}