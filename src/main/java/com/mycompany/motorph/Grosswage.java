/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.motorph;

import java.text.DecimalFormat;
import java.util.List;
import java.util.Scanner;

/**
 *
 * @author angeliquerivera
 */
public class Grosswage extends Calculation {
    public static String targetEmployeeID;
    public static String employeeName;
    private static String employeeID;
    public static int targetMonth;
    public static double gross;
    private static double hourly;
    public static double hours;

    // Declare decimalFormat
    private static final DecimalFormat decimalFormat = new DecimalFormat("#.##");

    public double calculate() {
        List<Employee> employees = EmployeeModelFromFile.getEmployeeModelList();  // Get the list of employees
        Scanner sc = new Scanner(System.in);

        // Asks for User Input.
        System.out.println("-------------------------");
        System.out.print("Enter Employee #: ");
        targetEmployeeID = sc.next();
        System.out.println("-------------------------");
        System.out.print("Enter Month: ");
        targetMonth = sc.nextInt();
        System.out.println("-------------------------");

        // Debugging print statements
        System.out.println("Target Employee ID: " + getTargetEmployeeID());
        System.out.println("Target Month: " + getTargetMonth());

        // Find the hourly rate for the chosen employee ID
        for (Employee employee : employees) {
            if (employee.getEmployeeNumber().equals(getTargetEmployeeID())) {
                setHourly(employee.getHourlyRate());
                employeeID = employee.getEmployeeNumber();

                // Check if the hourly rate string is a valid decimal number
                if (isValidDecimal(Double.toString(getHourly()))) {
                    double hourlyRate = employee.getHourlyRate();
                    long totalHours = 0;
                    List<Integer> years = List.of(2022, 2023, 2024);

                    // Loop through the years to calculate total hours
                    for (int year : years) {
                        totalHours += AttendanceRecord.calculateTotalHoursAndPrint(year, getTargetMonth(), getTargetEmployeeID());
                    }

                    // Debugging print statement for calculated hours
                    System.out.println("Total Calculated Hours: " + totalHours);
                    double hoursCalculated = hourlyRate * totalHours;

                    hours = totalHours;
                    gross = hoursCalculated;
                    employeeName = employee.getFirstName() + " " + employee.getLastName();

                    // Call printGross to display the results
                    printGross();
                } else {
                    System.out.println("Invalid hourly rate for Employee ID " + getEmployeeID() + ": " + getHourly());
                }
                return gross; // Exit the loop once the employee is found
            }
        }

        // If the loop completes without finding the employee ID
        System.out.println("Employee ID " + getEmployeeID() + " not found.");
        return gross;
    }

    public static void printGross() {
        System.out.println("""
                ------------------------------------------           
                Employee ID: %s
                Name: %s
                Hourly Rate: %.2f
                Total Hours: %s
                Gross Wage: %s
                ------------------------------------------
                """.formatted(getEmployeeID(),
                        employeeName, getHourly(),
                        hours,
                        decimalFormat.format(gross)
                ));
    }

    // CHECKS IF DECIMAL IN HOURLY RATE IS VALID
    private static boolean isValidDecimal(String str) {
        try {
            Double.parseDouble(str);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /**
     * @return the targetEmployeeID
     */
    public static String getTargetEmployeeID() {
        return targetEmployeeID;
    }

    /**
     * @return the targetMonth
     */
    public static int getTargetMonth() {
        return targetMonth;
    }

    /**
     * @return the employeeID
     */
    public static String getEmployeeID() {
        return employeeID;
    }

    /**
     * @return the hourly
     */
    public static double getHourly() {
        return hourly;
    }

    /**
     * @param aHourly the hourly to set
     */
    public static void setHourly(double aHourly) {
        hourly = aHourly;
    }
}