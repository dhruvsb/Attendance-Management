package attpackage;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.*;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Date;

/**
 * Created by Dhruv Sb on 15-10-2016.
 */

public class DateRange {

    public static void DateRangeMethod (LocalDate localDate1,LocalDate localDate2) throws IOException {
        {
            HSSFWorkbook WorkbookOut = new HSSFWorkbook();
            HSSFSheet SheetOut = WorkbookOut.createSheet("DateRange 1");
            FileInputStream FileInput = new FileInputStream(new File("C:\\CO-III.xls"));
            HSSFWorkbook workbookIn = new HSSFWorkbook(FileInput);
            HSSFSheet sheetIn = workbookIn.getSheetAt(0);
            HSSFRow row;
            HSSFCell cell;
            String[] name = new String[50];
            int[] enroll = new int[50];
            int[] AttendanceOverall = new int[51];
            
            for (int i = 0; i < 51; ++i)
                AttendanceOverall[i] = 0;
            for (int i = 0; i < 50; ++i) {
                row = sheetIn.getRow(i + 1);
                int enrollIn = (int) row.getCell(0).getNumericCellValue();
                enroll[i] = enrollIn;
                row = sheetIn.getRow(i + 1);
                String nameIn = row.getCell(1).getStringCellValue();
                name[i] = nameIn;
            }
            for (int i = 0; i < 50; ++i) {
                row = SheetOut.createRow(i + 1);
                cell = row.createCell(0);
                cell.setCellValue(enroll[i]);
                row = SheetOut.getRow(i + 1);
                cell = row.createCell(1);
                cell.setCellValue(name[i]);
            }
            for (int k = 1; k < 5; ++k) {
                HSSFSheet sheet = workbookIn.getSheetAt(k);
                int[] SubjectAttendance = new int[51];
                for (int i = 0; i < 51; ++i)
                    SubjectAttendance[i] = 0;
                for (int i = 1; i < 11; ++i) {
                    HSSFRow rowIn = sheet.getRow(i);
                    Date DateOriginal = rowIn.getCell(1).getDateCellValue();
                    LocalDate FormatDate = LocalDate.parse( new SimpleDateFormat("yyyy-MM-dd").format(DateOriginal) );
                    String DateInput = String.valueOf(FormatDate);
                    String StartDate = String.valueOf(localDate1);
                    String EndDate = String.valueOf(localDate2);
                    if (DateInput.compareTo(StartDate) >= 0 && DateInput.compareTo(EndDate) <= 0) {
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
                }
                for (int i = 0; i < 50; ++i) {
                    double present = (double) SubjectAttendance[i + 1];
                    double total = (double) SubjectAttendance[0];
                    double percentage = 100 * (1 - (present / total));
                    if (k == 0) {
                        row = SheetOut.getRow(i + 1);
                        cell = row.createCell(2);
                        String SubAttended = (SubjectAttendance[0] - SubjectAttendance[i + 1]) + "/" + SubjectAttendance[0];
                        cell.setCellValue(SubAttended);
                        cell = row.createCell(3);
                        cell.setCellValue(percentage);
                        
                    } else {
                        row = SheetOut.getRow(i + 1);
                        cell = row.createCell((k * 2));
                        String SubAttended = (SubjectAttendance[0] - SubjectAttendance[i + 1]) + "/" + SubjectAttendance[0];
                        cell.setCellValue(SubAttended);
                        cell = row.createCell((k * 2) + 1);
                        cell.setCellValue(percentage);
                        CellStyle style = WorkbookOut.createCellStyle();
                        if (percentage < 50) {
                            style.setFillForegroundColor(IndexedColors.RED.getIndex());
                            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                            cell.setCellStyle(style);
                        }
                    }
                }
            }
            for (int i = 1; i < 51; ++i) {
                row = SheetOut.getRow(i);
                cell = row.createCell(10);
                double present = (double) AttendanceOverall[i];
                double total = (double) AttendanceOverall[0];
                double percentage = 100 * (1 - (present / total));
                cell.setCellValue(percentage);
                CellStyle style = WorkbookOut.createCellStyle();
                if (percentage < 50) {
                    style.setFillForegroundColor(IndexedColors.RED.getIndex());
                    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                    cell.setCellStyle(style);
                } else if (percentage < 75) {
                    style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                    cell.setCellStyle(style);
                } else if (percentage > 95) {
                    style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                    cell.setCellStyle(style);
                }
            }
            try {
                FileOutputStream FileOutput = new FileOutputStream("DateRange.xls");
                WorkbookOut.write(FileOutput);
                FileOutput.close();
            } catch (Exception exc) {
                exc.printStackTrace();
            }
        }
    }
}