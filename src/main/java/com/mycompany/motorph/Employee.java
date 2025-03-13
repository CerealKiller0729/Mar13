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
public class Employee {
    private String employeeNumber;
    private String lastName;
    private String firstName;
    private String birthday;
    private String address;
    private String phoneNumber;
    private String sssNumber;
    private String philhealthNumber;
    private String tinNumber;
    private String pagIbigNumber;
    private String status;
    private String position;
    private String immediateSupervisor;
    private String basicSalary;
    private String riceSubsidy;
    private String phoneAllowance;
    private String clothingAllowance;
    private String grossSemiMonthlyRate;
    protected double hourlyRate;

    // Constructor to initialize an Employee object with data
    public Employee(String[] data) {
        this.employeeNumber = getValue(data, 0);
        // If the employee number is numeric in the Excel file, convert it to a string
    try {
        // Attempt to parse the employee number as a double
        double empNum = Double.parseDouble(this.employeeNumber);
        // Convert to an integer to remove decimal points and then to a string
        this.employeeNumber = String.valueOf((int) empNum);
    } catch (NumberFormatException e) {
        // If parsing fails, keep the original string
        System.err.println("Error parsing employee number: " + this.employeeNumber);
    }
        this.lastName = getValue(data, 1);
        this.firstName = getValue(data, 2);
        this.birthday = getValue(data, 3);
        this.address = getValue(data, 4);
        this.phoneNumber = getValue(data, 5);
        this.sssNumber = getValue(data, 6);
        this.philhealthNumber = getValue(data, 7);
        this.tinNumber = getValue(data, 8);
        this.pagIbigNumber = getValue(data, 9);
        this.status = getValue(data, 10);
        this.position = getValue(data, 11);
        this.immediateSupervisor = getValue(data, 12);
        this.basicSalary = getValue(data, 13);
        this.riceSubsidy = getValue(data, 14);
        this.phoneAllowance = getValue(data, 15);
        this.clothingAllowance = getValue(data, 16);
        this.grossSemiMonthlyRate = getValue(data, 17);
        this.hourlyRate = getDoubleValue(data, 18);
    }

    public Employee() {
    }

    // Method to load employees from an Excel file
    public static List<Employee> loadEmployees(String filePath) {
        List<Employee> employees = new ArrayList<>();

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
                    employees.add(new Employee(data));
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return employees;
    }

    // Overrides default toString to print employee data with spacing
    @Override
    public String toString() {
        String hourlyRateString = (Double.isNaN(hourlyRate)) ? "Hourly Rate" : Double.toString(hourlyRate);
        return String.format("%-15s%-15s%-15s%-15s%-80s%-15s%-15s%-15s%-20s%-15s%-15s%-35s%-25s%-15s%-15s%-20s%-20s%-25s%-15s",
                employeeNumber, lastName, firstName, birthday, address, phoneNumber, sssNumber, philhealthNumber, tinNumber,
                pagIbigNumber, status, position, immediateSupervisor, basicSalary, riceSubsidy, phoneAllowance, clothingAllowance,
                grossSemiMonthlyRate, hourlyRateString);
    }

    // Overloads toString to print in another format
    public String toString(boolean targetEmpTrue) {
        return String.format("""
                Employee ID: %s
                Name: %s %s
                Birthday: %s
                Address: %s
                PhoneNumber: %s
                SSSNumber: %s
                PhilHealth: %s
                TIN: %s
                PAG-IBIG: %s
                Status: %s
                Position: %s
                Supervisor: %s
                Basic Salary: %s
                Rice Subsidy: %s
                Phone allowance: %s
                Clothing Allowance: %s
                Gross Semi Monthly Rate: %s
                Hourly Rate: %s
                """,
                employeeNumber,
                lastName,
                firstName,
                birthday,
                address,
                phoneNumber,
                sssNumber,
                philhealthNumber,
                tinNumber,
                pagIbigNumber,
                status,
                position,
                immediateSupervisor,
                basicSalary,
                riceSubsidy,
                phoneAllowance,
                clothingAllowance,
                grossSemiMonthlyRate,
                hourlyRate
        );
    }

    private double getDoubleValue(String[] data, int index) {
        String value = getValue(data, index);

        // Check if the value is the header
        if (value.equalsIgnoreCase("Hourly Rate")) {
            return Double.NaN;
        }
          // Clean the value to remove any non-numeric characters
          value = value.replaceAll("[^0-9.]", ""); // Keep only digits and decimal points
         try {
            return Double.parseDouble(value);
        } catch (NumberFormatException e) {
            System.err.println("Error parsing double value at index " + index + " for Employee ID " + getEmployeeNumber() + ": " + value);
            return 0.0; // or another default value
        }
    }

    // Getters & Setters
    private String getValue(String[] data, int index) {
        return (data.length > index) ? data[index] : "";
    }

    // GETTER METHOD FOR EMPLOYEE NUMBER
    public String getEmployeeNumber() {
        return employeeNumber;
    }

    // Other getters and setters...

    public String getLastName() {
        return lastName;
    }

    public String getFirstName() {
        return firstName;
    }

    public String getBirthday() {
        return birthday;
    }

    public String getAddress() {
        return address;
    }

    public String getPhoneNumber() {
        return phoneNumber;
    }

    public String getSssNumber() {
        return sssNumber;
    }

    public String getPhilhealthNumber() {
        return philhealthNumber;
    }

    public String getTinNumber() {
        return tinNumber;
    }

    public String getPagIbigNumber() {
        return pagIbigNumber;
    }

    public String getStatus() {
        return status;
    }

    public String getPosition() {
        return position;
    }

    public String getImmediateSupervisor() {
        return immediateSupervisor;
    }

    public String getBasicSalary() {
        return basicSalary;
    }

    public String getRiceSubsidy() {
        return riceSubsidy;
    }

    public String getPhoneAllowance() {
        return phoneAllowance;
    }

    public String getClothingAllowance() {
        return clothingAllowance;
    }

    public String getGrossSemiMonthlyRate() {
        return grossSemiMonthlyRate;
    }

    public double getHourlyRate() {
        return hourlyRate;
    }

    public void setLastName(String lastName) {
        this.lastName = lastName;
    }

    public void setFirstName(String firstName) {
        this.firstName = firstName;
    }

    public void setHourlyRate(double hourlyRate) {
        this.hourlyRate = hourlyRate;
    }
}
