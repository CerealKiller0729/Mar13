/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.motorph;

import java.util.List;
import java.util.Scanner;

/**
 *
 * @author angeliquerivera
 */

public class MotorPHMain {
    private static Scanner scanner = new Scanner(System.in);

    public static void main(String[] args) {
        System.out.println("Current Working Directory: " + System.getProperty("user.dir"));

        // Load attendance records from the Excel file
        AttendanceRecord.loadAttendanceFromExcel("src/main/resources/AttendanceRecord.xlsx");
        menu();
    }

    // Print main menu
    private static void menu() {
        int resume = 1;
        do {
            System.out.print("""
                    ----- DASHBOARD-----

                    1: Show Employee Details
                    2: Calculate Gross Wage
                    3: Calculate Net Wage
                    0: EXIT
                    -------------------------
                    CHOOSE: """);

            String detailSub;
            String ch = scanner.next();

            switch (ch) {
                case "1":
                    System.out.print("""
                            ----- DASHBOARD-----

                            1: Individual Employee Details
 2: All Employee Details
                            -------------------------
                            Choose: """);
                    detailSub = scanner.next();
                    System.out.println("-------------------------");
                    menu(detailSub);
                    break;

                case "2":
                    Calculation grosswage = new Grosswage();
                    grosswage.calculate();
                    Grosswage.printGross();
                    break;

                case "3":
                    Calculation netwage = new Netwage();
                    netwage.calculate();
                    break;

                case "0":
                    System.exit(0);
                    break;

                default:
                    System.out.println("Invalid Input!");
                    break;
            }

            System.out.println("Back to menu? 1 = yes, 0 = no");
            resume = scanner.nextInt();
        } while (resume != 0);
    }

    // OVERLOAD MENU FOR SUBMENU IN PRINTING EMPLOYEE DETAILS
    private static void menu(String detailSub) {
        switch (detailSub) {
            case "1" -> printEmpSelectList();
            case "2" -> allEmployeeList();
        }
    }

    private static void printEmpSelectList() {
        List<Employee> employees = EmployeeModelFromFile.getEmployeeModelList();

        System.out.println("""
                -------------------------
                |     Employee List     |
                -------------------------""");

        String format = "%-15s%-20s"; // Adjust the width as needed

        for (Employee employee : employees) {
            System.out.printf(format, employee.getEmployeeNumber(), employee.getLastName());
            System.out.println(); // Print a new line
        }

        System.out.println("-------------------------");
        System.out.print("Enter Employee #: ");
        String empNum = scanner.next();
        System.out.println("-------------------------");

        for (Employee employee : employees) {
            if (employee.getEmployeeNumber().equals(empNum)) {
                System.out.println("Employee Details for Employee ID " + empNum + ":" + '\n' +
                        "-------------------------");
                System.out.println(employee.toString(true));
                System.out.println("-------------------------");
                return;
            }
        }

        System.out.println("Employee ID " + empNum + " not found.");
    }

    private static void allEmployeeList() {
        List<Employee> employees = EmployeeModelFromFile.getEmployeeModelList();

        for (Employee employee : employees) {
            System.out.println(employee);
        }

        System.out.println("-------------------------");
    }
}