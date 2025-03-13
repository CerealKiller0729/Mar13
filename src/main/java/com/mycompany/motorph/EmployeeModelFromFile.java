/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.motorph;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author angeliquerivera
 */
public class EmployeeModelFromFile {
    // Path to the Excel file containing Employee data
    private static String XLSX_FILE_PATH = "src/main/resources/Data.xlsx";
    private static final List<Employee> employees;

    // INITIALIZE
    static {
        employees = loadEmployees();
    }

    // LOADS EMPLOYEE DATA
    public static List<Employee> loadEmployees() {
        // INITIALIZES NEW OBJECT ARRAY LIST
        List<Employee> employees = new ArrayList<>();

        // TRIES TO OPEN AND LOAD Excel DATA THEN CLOSING RIGHT AFTER BY USING TRY-CATCH
        try (FileInputStream fis = new FileInputStream(XLSX_FILE_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            // Skip the header row
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    List<String> employeeData = new ArrayList<>();
                    for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                        Cell cell = row.getCell(cellIndex);
                        employeeData.add((cell != null) ? cell.toString() : ""); // Convert cell to string
                    }
                    if (employeeData.size() >= 19) { // Assuming you expect 19 fields
                        Employee employee = new Employee(employeeData.toArray(new String[0]));
                        employees.add(employee);
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return employees; // This returns the list of loaded Employee objects
    }

    /**
     * @return the employees
     */
    public static List<Employee> getEmployeeModelList() {
        return employees;
    }

    /** Set the path to the Excel file containing Employee data
     *  @param aXLSX_FILE_PATH the XLSX_FILE_PATH to set
     */
    public static void setXLSX_FILE_PATH(String aXLSX_FILE_PATH) {
        XLSX_FILE_PATH = aXLSX_FILE_PATH;
    }
}