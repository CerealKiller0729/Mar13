/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */

package com.mycompany.motorph;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;

public abstract class Calculation {

    protected static final DecimalFormat decimalFormat = new DecimalFormat("#.##");

    protected abstract double calculate();
}

class ExcelCalculation extends Calculation {

    private String filePath;

    public ExcelCalculation(String filePath) {
        this.filePath = filePath;
    }

    @Override
    protected double calculate() {
        double total = 0.0;

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            // Iterate through rows and perform calculations
            for (Row row : sheet) {
                // Assuming the first cell contains numeric values to sum
                Cell cell = row.getCell(0); // Change index based on your data structure
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    total += cell.getNumericCellValue();
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return total;
    }

    public String getFormattedTotal() {
        return decimalFormat.format(calculate());
    }
}