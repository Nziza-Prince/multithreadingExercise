package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

public class EmployeeExcelExporter {
    private static final Lock writeLock = new ReentrantLock();

    public void exportEmployeesToExcel(String filePath) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Employees");

        try (Connection connection = DatabaseConnection.getConnection();
             PreparedStatement statement = connection.prepareStatement("SELECT * FROM employees");
             ResultSet resultSet = statement.executeQuery()) {

            // Create header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("ID");
            headerRow.createCell(1).setCellValue("First Name");
            headerRow.createCell(2).setCellValue("Last Name");
            headerRow.createCell(3).setCellValue("Email");
            headerRow.createCell(4).setCellValue("Department");
            headerRow.createCell(5).setCellValue("Job Title");
            headerRow.createCell(6).setCellValue("Salary");
            headerRow.createCell(7).setCellValue("Hire Date");

            // Write data rows
            int rowNum = 1;
            while (resultSet.next()) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(resultSet.getInt("id"));
                row.createCell(1).setCellValue(resultSet.getString("first_name"));
                row.createCell(2).setCellValue(resultSet.getString("last_name"));
                row.createCell(3).setCellValue(resultSet.getString("email"));
                row.createCell(4).setCellValue(resultSet.getString("department"));
                row.createCell(5).setCellValue(resultSet.getString("job_title"));
                row.createCell(6).setCellValue(resultSet.getDouble("salary"));
                row.createCell(7).setCellValue(resultSet.getDate("hire_date").toString());
            }

            // Write to Excel file (thread-safe)
            writeLock.lock();
            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
                System.out.println("Excel file generated successfully");
            } finally {
                writeLock.unlock();
            }

        } catch (SQLException | IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
                System.out.println("Excel file generated successfully");
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}