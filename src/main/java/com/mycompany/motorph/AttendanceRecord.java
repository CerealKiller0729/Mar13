package com.mycompany.motorph;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;

/**
 *
 * @author angeliquerivera
 */

public class AttendanceRecord {

    private String name;
    private String id;
    private LocalDate date;
    private LocalTime timeIn;
    private LocalTime timeOut;
    private static String XLSX_FILE_PATH = "src/main/resources/AttendanceRecord.xlsx";

    public static ArrayList<AttendanceRecord> attendanceRecords = new ArrayList<>();
    private static final DateTimeFormatter timeFormatter = DateTimeFormatter.ofPattern("H:mm");
    private static final DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");

    // Constructor
    public AttendanceRecord(String name, String id, LocalDate date, LocalTime timeIn, LocalTime timeOut) {
        this.name = name;
        this.id = id;
        this.date = date;
        this.timeIn = timeIn;
        this.timeOut = timeOut;
    }

    // New constructor that accepts a String array
    public AttendanceRecord(String[] data) {
        if (data.length < 6) {
            throw new IllegalArgumentException("Insufficient data to create AttendanceRecord");
        }
        this.id = data[0];
        this.name = data[1] + " " + data[2].trim(); // Assuming name is split into two parts
        this.date = LocalDate.parse(data[3], dateFormatter);
        this.timeIn = LocalTime.parse(data[4], timeFormatter);
        this.timeOut = LocalTime.parse(data[5], timeFormatter);
    }

    public AttendanceRecord() {}

    // LOADS ATTENDANCE FROM AN XLSX FILE
    public static void loadAttendanceFromExcel(String filePath) {
        attendanceRecords = loadAttendance(filePath);
    }

    public static ArrayList<AttendanceRecord> loadAttendance(String filePath) {
        ArrayList<AttendanceRecord> attendanceRecords = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            // Skip the header row
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    String id = getCellValueAsString(row.getCell(0));
                    String name = getCellValueAsString(row.getCell(1));
                    String surname = getCellValueAsString(row.getCell(2)).trim();
                    LocalDate date = row.getCell(3).getLocalDateTimeCellValue().toLocalDate();
                    LocalTime timeIn = row.getCell(4).getLocalDateTimeCellValue().toLocalTime();
                    LocalTime timeOut = row.getCell(5).getLocalDateTimeCellValue().toLocalTime();

                    // Create a new AttendanceRecord using the existing constructor
                    attendanceRecords.add(new AttendanceRecord(name + " " + surname, id, date, timeIn, timeOut));
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return attendanceRecords;
    }

    // Helper method to get cell value as String
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    // CALCULATES HOURS PER ATTENDANCE RECORD
    private static long calculateHoursWorked(LocalTime timeIn, LocalTime timeOut) {
        Duration duration = Duration.between(timeIn, timeOut);
        return duration.toHours();
    }

    // TO CALCULATES HOURS WORKED ON A SPECIFIC MONTH OF AN EMPLOYEE
    public static long calculateTotalHoursAndPrint(int year, int month, String targetEmployeeId) {
        long totalHours = 0;
        String employeeName = "";
        System.out.println("Checking attendance for Employee ID: " + targetEmployeeId + " for Year: " + year + " Month: " + month);

        for (AttendanceRecord entry : attendanceRecords) {
            if (entry.getId().equals(targetEmployeeId)) {
                int entryYear = entry.getDate().getYear();
                int entryMonth = entry.getDate().getMonthValue();

                if (entryYear == year && entryMonth == month) {
                    long hoursWorked = calculateHoursWorked(entry.getTimeIn(), entry.getTimeOut());
                    totalHours += hoursWorked;
                    employeeName = entry.getName();
                }
            }
        }

        if (totalHours > 0) {
            System.out.printf("Employee ID: %s, Name: %s, Total Hours: %d%n", targetEmployeeId, employeeName, totalHours);
        } else {
            System.out.println("No hours found for Employee ID: " + targetEmployeeId);
        }

        return totalHours;
    }

    // Getters
    public String getName() {
        return name;
    }

    public String getId() {
        return id;
    }

    public LocalDate getDate() {
        return date;
    }

    public LocalTime getTimeIn() {
        return timeIn;
    }

    public LocalTime getTimeOut() {
        return timeOut;
    }

    public static ArrayList<AttendanceRecord> getAttendanceRecords() {
        return attendanceRecords;
    }
}