/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.motorph;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author angeliquerivera
 */
// This LatePenalty class extends Calculation and calculates the late penalty deduction for an employee.
public class LatePenalty extends Calculation {

    @Override
    public double calculate() {
        double totalLateDeduction = 0;
        boolean foundLateRecord = false; // Flag to check if any late records were found

        // Load attendance records from the Excel file
        List<AttendanceRecord> attendanceRecords = loadAttendanceFromExcel("src/main/resources/AttendanceRecord.xlsx");

        // Iterates through every attendance record.
        for (AttendanceRecord attendanceRecord : attendanceRecords) {
            // Check if the record is for the target employee.
            if (attendanceRecord.getId().equals(Grosswage.getTargetEmployeeID())) {
                LocalDate recordDate = attendanceRecord.getDate();
                int recordMonth = recordDate.getMonthValue(); // Month as an integer.

                // Check if the record is in the target month
                if (recordMonth == Grosswage.getTargetMonth()) {
                    // Assuming late penalty starts from 8:10 AM (490 minutes) onwards
                    final int lateThreshold = 490;

                    // Gets time in.
                    LocalTime timeIn = attendanceRecord.getTimeIn();

                    // Converts hour into minutes then add it to the minutes. Sample: 8:40 is 480 + 40.
                    int lateTime = timeIn.getHour() * 60 + timeIn.getMinute();

                    // Compares late time to threshold.
                    if (lateTime >= lateThreshold) {
                        // Calculate the per-minute equivalent of the hourly rate
                        double hourlyRate = Grosswage.getHourly();
                        double perMinuteRate = hourlyRate / 60.0;

                        // Calculate the deduction amount based on late time
                        double deduction = perMinuteRate * (lateTime - lateThreshold);

                        // Ensure deduction is non-negative
                        totalLateDeduction += Math.max(0, deduction);
                        // Set the flag to true since we found a late record
                        foundLateRecord = true;
                    }
                }
            }
        }

        // Print the target employee ID and month only if a late record was found
        if (foundLateRecord) {
            System.out.println("Late ID: " + Grosswage.getTargetEmployeeID());
            System.out.println("Month: " + Grosswage.getTargetMonth());
        }
        return totalLateDeduction;
    }

    // Method to load attendance records from an Excel file
    private List<AttendanceRecord> loadAttendanceFromExcel(String filePath) {
        List<AttendanceRecord> attendanceRecords = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            // Skip the header row
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    String[] data = new String[row.getLastCellNum()];
                    for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                        Cell cell = row.getCell(cellIndex);
                        data[cellIndex] = (cell != null) ? cell.toString() : ""; // Convert cell to string
                    }
                    attendanceRecords.add(new AttendanceRecord(data));
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return attendanceRecords;
    }
}